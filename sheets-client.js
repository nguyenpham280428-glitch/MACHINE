(function (global) {
  "use strict";

  const pending = new Map();
  const subscriptions = new Map();
  let frame = null;
  let bridgeWindow = null;
  let readyPromise = null;
  let passphrase = "";
  let nextId = 0;
  let pollTimer = null;
  let pollPending = false;

  function makeError(message) {
    return new Error(message || "Google Sheets bridge request failed");
  }

  function send(request) {
    return waitForFrame().then(
      () =>
        new Promise((resolve, reject) => {
          const id = "mmb-" + Date.now().toString(36) + "-" + (++nextId).toString(36);
          const timer = setTimeout(() => {
            pending.delete(id);
            reject(makeError("Apps Script không phản hồi. Kiểm tra URL và triển khai Web app."));
          }, 30000);
          pending.set(id, {
            resolve(value) {
              clearTimeout(timer);
              resolve(value);
            },
            reject(error) {
              clearTimeout(timer);
              reject(error);
            },
          });
          bridgeWindow.postMessage(
            { type: "mmb-call", id, request: Object.assign({ passphrase }, request) },
            "*"
          );
        })
    );
  }

  function waitForFrame() {
    if (readyPromise) return readyPromise;
    readyPromise = new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        readyPromise = null;
        reject(makeError("Không tải được Apps Script bridge trong 30 giây."));
      }, 30000);
      const onReady = (event) => {
        if (!isFrameDescendant(event.source) && !isAppsScriptOrigin(event.origin)) return;
        if (event.data && event.data.type === "mmb-ready") {
          bridgeWindow = event.source;
          event.source.postMessage({ type: "mmb-ack" }, "*");
          clearTimeout(timer);
          global.removeEventListener("message", onReady);
          resolve();
        }
      };
      global.addEventListener("message", onReady);
    });
    return readyPromise;
  }

  function isFrameDescendant(source) {
    if (!frame || !source) return false;
    let current = source;
    for (let depth = 0; depth < 10; depth++) {
      if (current === frame.contentWindow) return true;
      let parent;
      try {
        parent = current.parent;
      } catch (error) {
        return false;
      }
      if (!parent || parent === current) return false;
      current = parent;
    }
    return false;
  }

  function isAppsScriptOrigin(origin) {
    try {
      const url = new URL(origin);
      return (
        url.protocol === "https:" &&
        (url.hostname === "script.google.com" ||
          url.hostname === "script.googleusercontent.com" ||
          url.hostname.endsWith(".googleusercontent.com"))
      );
    } catch (error) {
      return false;
    }
  }

  function snapshotDoc(value) {
    return {
      exists: !!value.exists,
      id: value.id,
      data() {
        return value.data;
      },
    };
  }

  function snapshotQuery(value) {
    return {
      docs: (value.docs || []).map((doc) => ({
        id: doc.id,
        data() {
          return doc.data;
        },
      })),
    };
  }

  function notify(subscription, value) {
    const signature = JSON.stringify(value);
    if (subscription.signature === signature) return;
    subscription.signature = signature;
    subscription.errorSent = false;
    try {
      const result =
        subscription.kind === "doc"
          ? subscription.callback(snapshotDoc(value))
          : subscription.callback(snapshotQuery(value));
      if (result && typeof result.catch === "function") result.catch(subscription.error);
    } catch (error) {
      subscription.error(error);
    }
  }

  async function poll() {
    if (pollPending || subscriptions.size === 0) return;
    pollPending = true;
    const active = Array.from(subscriptions.values());
    try {
      const results = await send({
        action: "watch",
        targets: active.map(({ kind, path, collection, where, limit }) => ({
          kind,
          path,
          collection,
          where,
          limit,
        })),
      });
      active.forEach((subscription, index) => {
        if (subscriptions.has(subscription.id)) notify(subscription, results[index]);
      });
    } catch (error) {
      active.forEach((subscription) => {
        if (subscriptions.has(subscription.id) && !subscription.errorSent) {
          subscription.errorSent = true;
          subscription.error(error);
        }
      });
    } finally {
      pollPending = false;
    }
  }

  function subscribe(target, callback, onError) {
    const id = "sub-" + (++nextId).toString(36);
    const subscription = Object.assign({ id, callback, signature: null, errorSent: false }, target);
    subscription.error =
      typeof onError === "function" ? onError : (error) => global.console.error(error);
    subscriptions.set(id, subscription);
    if (!pollTimer) pollTimer = setInterval(poll, 20000);
    setTimeout(poll, 0);
    return () => {
      subscriptions.delete(id);
      if (!subscriptions.size && pollTimer) {
        clearInterval(pollTimer);
        pollTimer = null;
      }
    };
  }

  function createDb() {
    return {
      async setMany(documents) {
        for (let i = 0; i < documents.length; i += 100) {
          await send({ action: "setMany", documents: documents.slice(i, i + 100) });
        }
      },
      async deleteMany(paths) {
        for (let i = 0; i < paths.length; i += 100) {
          await send({ action: "deleteMany", paths: paths.slice(i, i + 100) });
        }
      },
      doc(path) {
        return {
          async get() {
            return snapshotDoc(await send({ action: "get", path }));
          },
          async set(data) {
            await send({ action: "set", path, data });
          },
          async delete() {
            await send({ action: "delete", path });
          },
          onSnapshot(callback, onError) {
            return subscribe({ kind: "doc", path }, callback, onError);
          },
        };
      },
      collection(name) {
        const makeQuery = (query) => ({
          where(field, operator, value) {
            return makeQuery(Object.assign({}, query, { where: { field, operator, value } }));
          },
          limit(count) {
            return makeQuery(Object.assign({}, query, { limit: count }));
          },
          async get() {
            return snapshotQuery(await send(Object.assign({ action: "query" }, query)));
          },
          onSnapshot(callback, onError) {
            return subscribe(Object.assign({ kind: "query" }, query), callback, onError);
          },
        });
        return makeQuery({ collection: name, where: null, limit: null });
      },
      close() {
        subscriptions.clear();
        if (pollTimer) clearInterval(pollTimer);
        pollTimer = null;
        if (frame) frame.remove();
        frame = null;
        readyPromise = null;
        bridgeWindow = null;
        passphrase = "";
      },
    };
  }

  function validUrl(value) {
    let url;
    try {
      url = new URL(value);
    } catch (error) {
      throw makeError("Dán URL Web app Apps Script có đuôi /exec.");
    }
    if (
      url.protocol !== "https:" ||
      url.hostname !== "script.google.com" ||
      !/^\/macros\/(?:u\/\d+\/)?s\/[^/]+\/exec\/?$/.test(url.pathname)
    ) {
      throw makeError("URL phải là Web app Apps Script dạng https://script.google.com/macros/s/.../exec.");
    }
    url.searchParams.set("bridge", "1");
    return url.toString();
  }

  function onMessage(event) {
    if (!frame || (event.source !== bridgeWindow && !isAppsScriptOrigin(event.origin))) return;
    const message = event.data;
    if (!message || message.type !== "mmb-result" || !pending.has(message.id)) return;
    const request = pending.get(message.id);
    pending.delete(message.id);
    if (message.ok) request.resolve(message.value);
    else request.reject(makeError(message.error));
  }

  const downloads = {
    async save(options) {
      const blob = options.data instanceof Blob ? options.data : new Blob([options.data]);
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = options.filename;
      link.style.display = "none";
      document.body.appendChild(link);
      link.click();
      link.remove();
      setTimeout(() => URL.revokeObjectURL(url), 60000);
    },
  };

  global.addEventListener("message", onMessage);

  global.MmbSheets = {
    downloads,
    async connect(url, secret) {
      if (frame) this.disconnect();
      passphrase = String(secret || "");
      if (passphrase.length < 16) {
        throw makeError("Mật khẩu dùng chung phải có ít nhất 16 ký tự.");
      }
      frame = document.createElement("iframe");
      frame.title = "Google Sheets connection";
      frame.setAttribute("aria-hidden", "true");
      frame.style.cssText =
        "position:fixed;width:1px;height:1px;left:-10px;bottom:0;border:0;opacity:0;pointer-events:none";
      frame.src = validUrl(url);
      document.body.appendChild(frame);
      readyPromise = null;
      const db = createDb();
      await db.doc("meta/crypto").get();
      return db;
    },
    disconnect() {
      subscriptions.clear();
      if (pollTimer) clearInterval(pollTimer);
      pollTimer = null;
      for (const request of pending.values()) request.reject(makeError("Đã ngắt kết nối."));
      pending.clear();
      if (frame) frame.remove();
      frame = null;
      readyPromise = null;
      bridgeWindow = null;
      passphrase = "";
    },
  };
})(window);

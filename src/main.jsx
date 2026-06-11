import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import App from "./App.jsx";
import "./index.css";

/* ---------------------------------------------------------------------------
   Local persistence shim.
   Inside Claude, an artifact gets a `window.storage` API automatically.
   On your own machine that API doesn't exist, so we provide a drop-in
   replacement backed by the browser's localStorage. Your data is then saved
   per-browser and survives refreshes. No other code changes are needed.
--------------------------------------------------------------------------- */
if (!window.storage) {
  window.storage = {
    async get(key) {
      const v = localStorage.getItem(key);
      return v === null ? null : { key, value: v };
    },
    async set(key, value) {
      localStorage.setItem(key, value);
      return { key, value };
    },
    async delete(key) {
      localStorage.removeItem(key);
      return { key, deleted: true };
    },
    async list(prefix = "") {
      const keys = Object.keys(localStorage).filter((k) => k.startsWith(prefix));
      return { keys, prefix };
    },
  };
}

createRoot(document.getElementById("root")).render(
  <StrictMode>
    <App />
  </StrictMode>
);

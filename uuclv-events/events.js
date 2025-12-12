/*
 * UUCLV Website Event Feed Embed (iframe-safe, copy-paste snippet)
 *
 * This file is meant to be opened and COPY–PASTED into an <script> tag
 * inside an iframe (or any HTML page) – it is NOT loaded via <script src>.
 *
 * Usage (high level):
 *   1. Create a container element in your HTML, e.g. <div id="uuclv-events-all-upcoming"></div>
 *   2. Add a <script> block and paste this entire file into it.
 *   3. Edit the CONFIG object below for that iframe:
 *      - sheetId or sheetUrl (one of them must be provided)
 *      - sheetTab (e.g. "All Upcoming")
 *      - targetSelector (e.g. "#uuclv-events-all-upcoming")
 *   4. Optionally adjust other options (styles, messages) if desired.
 *
 * You can have multiple iframes on the same page, each with its own
 * copy of this script and its own CONFIG, pointing at different sheets
 * or tabs. Each iframe runs in an isolated JS environment, so there is
 * no cross-talk between embeds.
 */

(function () {
  "use strict";

  // ==========================
  // CONFIG – EDIT PER IFRAME
  // ==========================
  var CONFIG = {
    // Either provide sheetId OR sheetUrl. If both are set, sheetId wins.
    //
    // Example sheet URL:
    //   https://docs.google.com/spreadsheets/d/10uLUdC-eDJlkL_hF0Mj20Gr2i-qXKhQSBozhVeVl-XE/edit#gid=0
    sheetId: "",          // <-- put your Sheet ID here (recommended)
    sheetUrl: "",         // <-- or put a full Sheet URL here instead

    // Exact tab name in the sheet, e.g. "All Upcoming" or a filtered view tab.
    sheetTab: "All Upcoming",

    // CSS selector for the container where the event list should be rendered.
    // Typically this is an ID, e.g. "#uuclv-events-all-upcoming".
    targetSelector: "#uuclv-public-events",

    // If true, injects a small, modern default stylesheet for the cards.
    // Set to false if you prefer to provide your own styles.
    injectDefaultStyles: true,

    // Text messages shown to users in various error/empty cases.
    messages: {
      unsupportedBrowser: "This events list cannot be shown because your web browser is too old or does not support the features this page needs.",
      loadError: "We couldn\'t load the events right now. Please try again later or contact the office.",
      noEvents: "There are no upcoming events to show right now.",
      badConfig: "Events could not be loaded because the event feed is not configured correctly.",
    }
  };

  // ==========================
  // IMPLEMENTATION
  // ==========================

  function main() {
    var target = selectTarget(CONFIG.targetSelector);
    if (!target) {
      // No suitable target; nothing to show. Fail silently for end users.
      return;
    }

    if (!supportsRequiredFeatures()) {
      renderMessage(target, CONFIG.messages.unsupportedBrowser, "uuclv-ev-error");
      return;
    }

    if (CONFIG.injectDefaultStyles) {
      injectDefaultStyles();
    }

    target.className += (target.className ? " " : "") + "uuclv-ev-root";

    var sheetId = CONFIG.sheetId || extractSheetIdFromUrl(CONFIG.sheetUrl || "");
    if (!sheetId) {
      console.error("UUCLV events embed: missing or invalid sheetId/sheetUrl in CONFIG.");
      renderMessage(target, CONFIG.messages.badConfig, "uuclv-ev-error");
      return;
    }

    var tabName = CONFIG.sheetTab || "All Upcoming";

    var url = "https://docs.google.com/spreadsheets/d/" + encodeURIComponent(sheetId) +
              "/gviz/tq?tqx=out:json&headers=1&sheet=" + encodeURIComponent(tabName);

    fetch(url)
      .then(function (response) { return response.text(); })
      .then(function (text) {
        var data;
        try {
          data = parseGvizJson(text);
        } catch (err) {
          console.error("UUCLV events embed: could not parse sheet response", err);
          renderMessage(target, CONFIG.messages.loadError, "uuclv-ev-error");
          return;
        }

        if (!data || !data.table || !data.table.rows) {
          console.error("UUCLV events embed: malformed table data");
          renderMessage(target, CONFIG.messages.loadError, "uuclv-ev-error");
          return;
        }

        var rows = data.table.rows || [];
        if (!rows.length) {
          renderMessage(target, CONFIG.messages.noEvents, "uuclv-ev-empty");
          return;
        }

        // For the default All Upcoming view, the column order is:
        // 0 Approver | 1 Event Name | 2 Description | 3 Start | 4 End |
        // 5 Target Audience | 6 Building Spaces | 7 Advertise Where | ...
        var html = buildEventsHtmlFromAllUpcoming(rows);
        target.innerHTML = html;

        attachToggleHandler(target);
      })
      .catch(function (err) {
        console.error("UUCLV events embed: fetch error", err);
        renderMessage(target, CONFIG.messages.loadError, "uuclv-ev-error");
      });
  }

  // --------------------------
  // Environment + DOM helpers
  // --------------------------

  function supportsRequiredFeatures() {
    try {
      return !!(window && window.fetch && window.Promise && document &&
                document.querySelector && window.JSON && JSON.parse &&
                Element && Element.prototype && Element.prototype.addEventListener);
    } catch (e) {
      return false;
    }
  }

  function selectTarget(selector) {
    if (!document || !selector) return null;
    try {
      return document.querySelector(selector);
    } catch (e) {
      return null;
    }
  }

  function renderMessage(target, text, extraClass) {
    var cls = "uuclv-ev-message" + (extraClass ? " " + extraClass : "");
    target.innerHTML = "<p class=\"" + cls + "\">" + escapeHtml(text || "") + "</p>";
  }

  function injectDefaultStyles() {
    if (!document || !document.head) return;

    // Avoid injecting multiple times in the same iframe
    if (document.getElementById("uuclv-ev-styles")) {
      return;
    }

    var css = "" +
      ".uuclv-ev-root{font-family:system-ui,-apple-system,Segoe UI,Roboto,Arial,sans-serif;max-width:900px;margin:0 auto;font-size:14px;line-height:1.45;}" +
      ".uuclv-ev-card{border:1px solid #e5e7eb;border-radius:14px;padding:14px 16px;margin:12px 0;box-shadow:0 1px 2px rgba(0,0,0,.04);background:#ffffff;}" +
      ".uuclv-ev-title{margin:0 0 4px;font-size:1.05rem;}" +
      ".uuclv-ev-title-link{background:none;border:0;padding:0;margin:0;font:inherit;color:#0000EE;text-decoration:underline;cursor:pointer;text-align:left;}" +
      ".uuclv-ev-title-link:hover{text-decoration:none;}" +
      ".uuclv-ev-title-link:focus-visible{outline:2px solid #2563eb;outline-offset:2px;border-radius:4px;}" +
      ".uuclv-ev-meta{font-size:.9rem;margin:4px 0;color:#374151;}" +
      ".uuclv-ev-meta strong{font-weight:600;}" +
      ".uuclv-ev-desc{margin-top:8px;white-space:pre-wrap;display:none;color:#111827;}" +
      ".uuclv-ev-card--open .uuclv-ev-desc{display:block;}" +
      ".uuclv-ev-empty,.uuclv-ev-error,.uuclv-ev-message{text-align:center;color:#6b7280;margin:24px 0;font-size:.95rem;}" +
      ".uuclv-ev-error{color:#b91c1c;}";

    var styleEl = document.createElement("style");
    styleEl.id = "uuclv-ev-styles";
    styleEl.type = "text/css";
    if (styleEl.styleSheet) {
      styleEl.styleSheet.cssText = css;
    } else {
      styleEl.appendChild(document.createTextNode(css));
    }
    document.head.appendChild(styleEl);
  }

  // --------------------------
  // Data parsing + rendering
  // --------------------------

  function extractSheetIdFromUrl(url) {
    if (!url) return "";
    var m = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    return m && m[1] ? m[1] : "";
  }

  function parseGvizJson(text) {
    // The gviz endpoint returns something like: "/*O_o*/\ngoogle.visualization.Query.setResponse({...});"
    // We strip everything before the first "{" and after the last "}".
    var firstBrace = text.indexOf("{");
    var lastBrace = text.lastIndexOf("}");
    if (firstBrace === -1 || lastBrace === -1 || lastBrace <= firstBrace) {
      throw new Error("Invalid gviz response format");
    }
    var jsonString = text.substring(firstBrace, lastBrace + 1);
    return JSON.parse(jsonString);
  }

  function cellValue(row, index) {
    var c = row && row.c && row.c[index];
    if (!c) return "";
    if (c.f != null && c.f !== "") return String(c.f);
    if (c.v != null && c.v !== "") return String(c.v);
    return "";
  }

  function buildEventsHtmlFromAllUpcoming(rows) {
    var out = [];
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      if (!r || !r.c) continue;

      var title = cellValue(r, 1);       // Event Name
      var desc = cellValue(r, 2);        // Description
      var startText = cellValue(r, 3);   // Start (formatted in sheet)
      var endText = cellValue(r, 4);     // End (formatted in sheet)
      var location = cellValue(r, 6);    // Building Spaces as "Where"

      if (!title && !startText && !desc) {
        continue; // skip completely empty rows
      }

      var whenLine = "";
      if (startText) {
        whenLine = escapeHtml(startText);
        if (endText) {
          var endTime = extractTimePart(endText);
          if (endTime && endTime !== startText) {
            whenLine += " \u2192 " + escapeHtml(endTime); // arrow, end time only
          }
        }
      }

      var parts = [];
      parts.push("<article class=\"uuclv-ev-card\">");

      // Title with toggle link
      parts.push("<h3 class=\"uuclv-ev-title\">");
      parts.push("<button type=\"button\" class=\"uuclv-ev-title-link\" data-ev-toggle=\"1\" aria-expanded=\"false\">");
      parts.push(escapeHtml(title || "Untitled event"));
      parts.push("</button>");
      parts.push("</h3>");

      // When
      if (whenLine) {
        parts.push("<div class=\"uuclv-ev-meta\"><strong>When:</strong> " + whenLine + "</div>");
      }

      // Where
      if (location) {
        parts.push("<div class=\"uuclv-ev-meta\"><strong>Where:</strong> " + escapeHtml(location) + "</div>");
      }

      // Description (collapsible)
      if (desc) {
        parts.push("<div class=\"uuclv-ev-desc\">" + linkifyText(desc) + "</div>");
      }

      parts.push("</article>");
      out.push(parts.join(""));
    }

    if (!out.length) {
      return "<p class=\"uuclv-ev-empty\">" + escapeHtml(CONFIG.messages.noEvents) + "</p>";
    }

    return out.join("");
  }

  // --------------------------
  // Expand/collapse behavior
  // --------------------------

  function attachToggleHandler(root) {
    if (!root || !root.addEventListener) return;

    root.addEventListener("click", function (event) {
      var target = event.target || event.srcElement;
      var toggleEl = findAncestorWithAttribute(target, "data-ev-toggle");
      if (!toggleEl) return;

      event.preventDefault && event.preventDefault();

      var card = findAncestorWithClass(toggleEl, "uuclv-ev-card");
      if (!card) return;

      var isOpen = hasClass(card, "uuclv-ev-card--open");
      if (isOpen) {
        removeClass(card, "uuclv-ev-card--open");
        toggleEl.setAttribute("aria-expanded", "false");
      } else {
        addClass(card, "uuclv-ev-card--open");
        toggleEl.setAttribute("aria-expanded", "true");
      }
    });
  }

  function findAncestorWithAttribute(el, attr) {
    while (el && el !== document) {
      if (el.getAttribute && el.getAttribute(attr) != null) {
        return el;
      }
      el = el.parentNode;
    }
    return null;
  }

  function findAncestorWithClass(el, cls) {
    while (el && el !== document) {
      if (hasClass(el, cls)) return el;
      el = el.parentNode;
    }
    return null;
  }

  function hasClass(el, cls) {
    if (!el || !cls) return false;
    if (el.classList && el.classList.contains) {
      return el.classList.contains(cls);
    }
    var classes = (el.className || "").split(/\s+/);
    for (var i = 0; i < classes.length; i++) {
      if (classes[i] === cls) return true;
    }
    return false;
  }

  function addClass(el, cls) {
    if (!el || !cls) return;
    if (el.classList && el.classList.add) {
      el.classList.add(cls);
      return;
    }
    if (!hasClass(el, cls)) {
      el.className = (el.className ? el.className + " " : "") + cls;
    }
  }

  function removeClass(el, cls) {
    if (!el || !cls) return;
    if (el.classList && el.classList.remove) {
      el.classList.remove(cls);
      return;
    }
    var classes = (el.className || "").split(/\s+/);
    var out = [];
    for (var i = 0; i < classes.length; i++) {
      if (classes[i] && classes[i] !== cls) out.push(classes[i]);
    }
    el.className = out.join(" ");
  }

  // --------------------------
  // Text helpers
  // --------------------------

  function escapeHtml(s) {
    return String(s == null ? "" : s).replace(/[&<>"']/g, function (m) {
      return ({
        "&": "&amp;",
        "<": "&lt;",
        ">": "&gt;",
        "\"": "&quot;",
        "'": "&#39;"
      })[m] || m;
    });
  }

  function linkifyText(text) {
    if (!text) return "";
    var s = String(text);

    // Match either a URL or an email address.
    // - URLs: http:// or https:// followed by non-space characters
    // - Emails: simple local@domain pattern
    var tokenRe = /(https?:\/\/[^\s]+)|([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/g;

    var out = "";
    var lastIndex = 0;
    var match;

    while ((match = tokenRe.exec(s)) !== null) {
      var full = match[0];
      out += escapeHtml(s.slice(lastIndex, match.index));

      // Trim common trailing punctuation from the clickable part, but keep it in the text.
      var core = full;
      var trailing = "";
      var mTrail = full.match(/([.,!?;:)]*)$/);
      if (mTrail && mTrail[1]) {
        var t = mTrail[1];
        core = full.slice(0, full.length - t.length);
        trailing = t;
      }

      if (match[1]) {
        // URL
        var url = core;
        out += "<a href=\"" + escapeHtml(url) + "\" target=\"_blank\" rel=\"noopener noreferrer\">" +
               escapeHtml(url) + "</a>" + escapeHtml(trailing);
      } else if (match[2]) {
        // Email
        var email = core;
        out += "<a href=\"mailto:" + escapeHtml(email) + "\">" + escapeHtml(email) + "</a>" +
               escapeHtml(trailing);
      } else {
        out += escapeHtml(full);
      }

      lastIndex = tokenRe.lastIndex;
    }

    out += escapeHtml(s.slice(lastIndex));
    return out;
  }

  function extractTimePart(s) {
    if (!s) return "";
    var str = String(s);
    // Try to find a time-like fragment such as "6:00 PM", "18:30", "6:00:00 PM", etc.
    // - hours: 1–2 digits
    // - minutes: 2 digits
    // - optional seconds: :SS
    // - optional AM/PM (case-insensitive) with optional spaces (incl. narrow no-break)
    var m = str.match(/(\d{1,2}:\d{2}(?::\d{2})?\s*[APap]?[Mm]?)/);
    if (m && m[1]) {
      // Normalize internal spacing (e.g., "6:00   PM" -> "6:00 PM")
      return m[1].replace(/\s+/g, " ").trim();
    }
    // Fallback: if we can't detect a time, just return the original string.
    return str;
  }

  // Run immediately on script load
  try {
    main();
  } catch (e) {
    try {
      var target = selectTarget(CONFIG && CONFIG.targetSelector);
      if (target) {
        renderMessage(target, CONFIG.messages.loadError, "uuclv-ev-error");
      }
    } catch (ignore) {
      // swallow secondary errors; nothing else we can do in this environment
    }
    console.error("UUCLV events embed: unexpected error during initialization", e);
  }

})();

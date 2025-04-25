// ==UserScript==
// @name         Rotate_website_for_ical_to_ganto
// @namespace    http://tampermonkey.net/
// @version      2025-04-25
// @description  try to take over the world!
// @author       You
// @match        http://saclaopr19.spring8.or.jp/~lognote/calendar/gantt-group-tasks-together.html
// @icon         https://www.google.com/s2/favicons?sz=64&domain=tampermonkey.net
// @grant        none
// ==/UserScript==

(function() {
    'use strict';
    const b = document.body.style;
    b.transform = "rotate(90deg) scale(1.2)";
    b.position = "absolute";
    b.top = "370";
    b.left = "-200";
})();
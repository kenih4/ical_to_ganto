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
    //ページを回転
    const b = document.body.style;
    b.transform = "rotate(90deg) scale(1.2)";
    b.position = "absolute";
    b.top = "370";
    b.left = "-200";

    //時計表示
    const div = document.createElement('div');
    div.id = 'custom-clock';
    div.style.position = 'fixed';
    div.style.top = '750px';
    div.style.right = '1260px';
    div.style.backgroundColor = 'rgba(255, 255, 255, 0.1)';
    div.style.color ='rgba(255, 255, 1, 0.8)';
    div.style.padding = '10px';
    div.style.fontSize = '16px';
    div.style.fontFamily = 'monospace';
    div.style.zIndex = 10000;
    div.style.transform = 'rotate(270deg)';
    document.body.appendChild(div);

    function updateClock() {
        const now = new Date();
        div.textContent = now.toLocaleString('ja-JP');
    }
    updateClock();
    setInterval(updateClock, 1500);

})();
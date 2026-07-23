import { useState, useEffect } from "react";
import { supabase } from "./supabaseClient";
import * as XLSX from "xlsx";
import { ChevronLeft, Plus, Minus, Lock, ClipboardList, BarChart3, TrendingUp, Printer, ShieldAlert, ChevronRight, X, Settings, Trash2, UserPlus, Building2, Upload, Download, FileSpreadsheet, Briefcase, Image, FileText, CheckCircle2, Circle, Clock, ChevronDown, CalendarDays, Pencil } from "lucide-react";

// ---------------- BRAND (ActiveMe 360 guidelines) ----------------
const BC = {
  ink: "#2E054D", purple: "#4A1671", mid: "#944FC9", lilac: "#C77EFF",
  bg: "#F6F3F8", lime: "#E5FF00", white: "#FFFFFF",
};
const RCOL = { E: "#FFAC38", D: "#47ABFB", Es: "#189E8A" };
const RTXT = { E: BC.ink, D: BC.ink, Es: "#FFFFFF" };
const pill = (bg, fg) => ({ backgroundColor: bg, color: fg });
const tintBox = h => ({ backgroundColor: h + "1F", borderColor: h });

const FONTS = (
  <style>{"@import url('https://fonts.googleapis.com/css2?family=Public+Sans:wght@400;600;700;800&family=Unbounded:wght@500;700&display=swap'); .fd{font-family:'Unbounded',ui-sans-serif,system-ui;letter-spacing:-0.01em;} .fb{font-family:'Public Sans',ui-sans-serif,system-ui;}"}</style>
);

// Real ActiveMe 360 logo mark, extracted from brand artwork (rings + wordmark as vector paths).
const AM_RINGS = [
  { c: "#63DFC0", d: "M 205.804688 129.304688 C 200 131.34375 196.953125 137.703125 198.988281 143.503906 C 201.027344 149.304688 207.386719 152.355469 213.1875 150.316406 C 298.257812 120.421875 391.792969 165.308594 421.6875 250.382812 C 451.585938 335.453125 406.695312 428.988281 321.621094 458.882812 C 236.550781 488.777344 143.019531 443.890625 113.121094 358.816406 C 111.082031 353.015625 104.726562 349.964844 98.925781 352.003906 C 93.121094 354.042969 90.070312 360.398438 92.109375 366.203125 C 126.078125 462.859375 232.347656 513.863281 329.007812 479.894531 C 425.664062 445.929688 476.667969 339.65625 442.699219 243 C 408.734375 146.339844 302.460938 95.339844 205.804688 129.304688" },
  { c: "#2554C7", d: "M 192.625 91.804688 C 135.785156 111.78125 90.125 152.695312 64.058594 207.007812 C 61.398438 212.554688 63.734375 219.207031 69.277344 221.867188 C 74.824219 224.527344 81.476562 222.1875 84.136719 216.644531 C 115.203125 151.914062 176.292969 111.101562 242.757812 102.875 C 280.050781 98.257812 319.039062 103.898438 355.359375 121.332031 C 456.414062 169.832031 499.171875 291.5 450.671875 392.554688 C 427.179688 441.507812 386.027344 478.382812 334.800781 496.386719 C 283.574219 514.386719 228.402344 511.363281 179.449219 487.867188 C 173.90625 485.207031 167.253906 487.546875 164.59375 493.089844 C 161.929688 498.636719 164.269531 505.285156 169.8125 507.945312 C 224.128906 534.015625 285.34375 537.371094 342.183594 517.394531 C 399.023438 497.421875 444.683594 456.507812 470.75 402.191406 C 496.820312 347.875 500.175781 286.660156 480.199219 229.820312 C 460.226562 172.980469 419.3125 127.320312 364.996094 101.253906 C 310.679688 75.1875 249.464844 71.832031 192.625 91.804688" },
  { c: "#E6235C", d: "M 178.648438 52.027344 C 151.289062 61.644531 125.90625 75.449219 103.207031 93.066406 C 98.347656 96.835938 97.464844 103.832031 101.234375 108.691406 C 105.003906 113.550781 112 114.429688 116.859375 110.660156 C 137.664062 94.515625 160.9375 81.859375 186.03125 73.039062 C 313.714844 28.167969 454.097656 95.542969 498.964844 223.226562 C 543.835938 350.910156 476.460938 491.292969 348.78125 536.164062 C 342.976562 538.199219 339.929688 544.558594 341.964844 550.359375 C 344.003906 556.160156 350.363281 559.210938 356.164062 557.171875 C 423.628906 533.464844 477.820312 484.902344 508.761719 420.433594 C 539.703125 355.964844 543.683594 283.308594 519.976562 215.84375 C 496.269531 148.378906 447.707031 94.183594 383.238281 63.242188 C 318.769531 32.304688 246.109375 28.320312 178.648438 52.027344" },
];
const AM_WORDMARK = [
  "M 346.542969 255.875 C 344.914062 247.070312 338.714844 242.5 330.992188 242.5 C 322.292969 242.5 316.199219 247.609375 314.570312 255.875 Z M 314.675781 265.449219 C 316.417969 273.710938 323.050781 278.28125 332.074219 278.28125 C 338.277344 278.28125 342.953125 275.996094 346.105469 271.539062 L 355.789062 277.195312 C 350.460938 285.132812 342.519531 289.046875 331.972656 289.046875 C 323.160156 289.046875 315.984375 286.328125 310.542969 281 C 305.109375 275.5625 302.386719 268.710938 302.386719 260.441406 C 302.386719 252.289062 305.109375 245.546875 310.433594 240.109375 C 315.765625 234.558594 322.726562 231.839844 331.101562 231.839844 C 339.042969 231.839844 345.570312 234.667969 350.675781 240.21875 C 355.789062 245.761719 358.398438 252.617188 358.398438 260.554688 C 358.398438 261.75 358.292969 263.382812 357.964844 265.449219 Z M 314.675781 265.449219",
  "M 209.761719 336.367188 C 221.011719 340.28125 228.59375 349.382812 228.59375 361.136719 C 228.59375 370.363281 225.433594 377.691406 219.242188 382.871094 C 213.050781 388.054688 205.339844 390.707031 196.367188 390.707031 C 182.339844 390.707031 170.585938 384.261719 165.40625 371.75 L 182.84375 361.640625 C 184.992188 367.832031 189.542969 370.992188 196.367188 370.992188 C 212.292969 372.003906 212.292969 350.265625 196.367188 351.277344 L 192.199219 351.277344 L 184.488281 339.777344 L 200.539062 319.429688 L 168.183594 319.429688 L 168.183594 300.472656 L 225.054688 300.472656 L 225.054688 316.902344 Z M 209.761719 336.367188",
  "M 276.738281 368.464844 C 279.140625 366.191406 280.40625 363.15625 280.40625 359.492188 C 280.40625 355.828125 279.140625 352.796875 276.738281 350.644531 C 271.808594 346.097656 263.722656 346.097656 258.792969 350.519531 C 256.394531 352.667969 255.128906 355.703125 255.128906 359.492188 C 255.128906 363.285156 256.394531 366.316406 258.792969 368.464844 C 263.722656 372.890625 271.808594 372.890625 276.738281 368.464844 M 271.179688 328.402344 C 279.898438 329.035156 286.976562 332.320312 292.410156 338.007812 C 297.84375 343.695312 300.625 350.898438 300.625 359.492188 C 300.625 368.71875 297.59375 376.175781 291.398438 381.988281 C 285.332031 387.800781 277.496094 390.707031 267.765625 390.707031 C 258.035156 390.707031 250.074219 387.800781 244.007812 381.988281 C 237.941406 376.175781 234.910156 368.71875 234.910156 359.492188 C 234.910156 352.667969 236.679688 346.726562 240.089844 341.546875 L 267.136719 300.472656 L 290.386719 300.472656 Z M 271.179688 328.402344",
  "M 342.957031 370.992188 C 353.449219 370.992188 358.753906 362.273438 358.753906 344.707031 C 358.753906 327.140625 353.449219 318.417969 342.957031 318.417969 C 332.46875 318.417969 327.164062 327.140625 327.164062 344.707031 C 327.164062 362.273438 332.46875 370.992188 342.957031 370.992188 M 342.957031 390.707031 C 331.710938 390.707031 322.863281 386.539062 316.421875 378.195312 C 310.101562 369.726562 306.941406 358.609375 306.941406 344.707031 C 306.941406 330.804688 310.101562 319.683594 316.421875 311.34375 C 322.863281 302.875 331.710938 298.703125 342.957031 298.703125 C 354.207031 298.703125 363.054688 302.875 369.371094 311.34375 C 375.816406 319.683594 378.976562 330.804688 378.976562 344.707031 C 378.976562 358.609375 375.816406 369.726562 369.371094 378.195312 C 363.054688 386.539062 354.207031 390.707031 342.957031 390.707031",
  "M 82.886719 277.847656 C 87.78125 277.847656 91.917969 276.214844 95.179688 272.953125 C 98.441406 269.582031 100.070312 265.449219 100.070312 260.441406 C 100.070312 255.441406 98.441406 251.308594 95.179688 248.042969 C 91.917969 244.675781 87.78125 243.042969 82.886719 243.042969 C 77.996094 243.042969 73.859375 244.675781 70.597656 248.042969 C 67.339844 251.308594 65.707031 255.441406 65.707031 260.441406 C 65.707031 265.449219 67.339844 269.582031 70.597656 272.953125 C 73.859375 276.214844 77.996094 277.847656 82.886719 277.847656 M 100.070312 233.253906 L 111.820312 233.253906 L 111.820312 287.632812 L 100.070312 287.632812 L 100.070312 279.804688 C 95.609375 286.003906 89.304688 289.046875 81.039062 289.046875 C 73.539062 289.046875 67.226562 286.328125 61.894531 280.78125 C 56.566406 275.234375 53.953125 268.382812 53.953125 260.441406 C 53.953125 252.398438 56.566406 245.65625 61.894531 240.109375 C 67.226562 234.558594 73.539062 231.839844 81.039062 231.839844 C 89.304688 231.839844 95.609375 234.882812 100.070312 240.976562 Z M 100.070312 233.253906",
  "M 130.636719 280.78125 C 125.203125 275.234375 122.480469 268.492188 122.480469 260.441406 C 122.480469 252.289062 125.203125 245.546875 130.636719 240.109375 C 136.1875 234.558594 143.035156 231.839844 151.195312 231.839844 C 161.746094 231.839844 171.097656 237.277344 175.445312 245.871094 L 165.335938 251.746094 C 162.832031 246.519531 157.613281 243.367188 151.085938 243.367188 C 146.300781 243.367188 142.277344 245 139.015625 248.261719 C 135.859375 251.527344 134.230469 255.550781 134.230469 260.441406 C 134.230469 265.339844 135.859375 269.363281 139.015625 272.628906 C 142.277344 275.886719 146.300781 277.519531 151.085938 277.519531 C 157.503906 277.519531 162.941406 274.261719 165.554688 269.144531 L 175.664062 274.910156 C 170.988281 283.609375 161.746094 289.046875 151.195312 289.046875 C 143.035156 289.046875 136.1875 286.332031 130.636719 280.78125",
  "M 225.160156 233.257812 L 236.90625 233.257812 L 236.90625 287.632812 L 225.160156 287.632812 Z M 225.703125 222.921875 C 222.761719 219.875 222.761719 215.308594 225.703125 212.375 C 228.640625 209.328125 233.425781 209.328125 236.359375 212.375 C 239.300781 215.308594 239.300781 219.875 236.257812 222.921875 C 233.316406 225.855469 228.75 225.855469 225.703125 222.921875",
  "M 287.375 233.257812 L 300.207031 233.257812 L 278.996094 287.632812 L 265.402344 287.632812 L 244.199219 233.257812 L 257.027344 233.257812 L 272.253906 274.582031 Z M 287.375 233.257812",
  "M 189.269531 289.589844 L 201.011719 289.589844 L 201.011719 244.566406 L 214.5 244.566406 L 214.5 233.257812 L 201.011719 233.257812 L 201.011719 218.027344 L 189.269531 221.507812 L 189.269531 233.257812 L 179.257812 233.257812 L 179.257812 244.566406 L 189.269531 244.566406 Z M 189.269531 289.589844",
  "M 111.953125 315.671875 L 111.953125 338.957031 L 103.839844 338.957031 L 103.839844 316.195312 C 103.839844 311.011719 101.058594 308.007812 96.550781 308.007812 C 91.519531 308.007812 88.363281 311.3125 88.363281 317.925781 L 88.363281 338.957031 L 80.25 338.957031 L 80.25 316.195312 C 80.25 311.011719 77.695312 308.007812 73.1875 308.007812 C 68.304688 308.007812 64.851562 311.386719 64.851562 317.925781 L 64.851562 338.957031 L 56.738281 338.957031 L 56.738281 301.398438 L 64.851562 301.398438 L 64.851562 305.90625 C 67.253906 302.226562 70.859375 300.421875 75.59375 300.421875 C 80.398438 300.421875 83.929688 302.375 86.261719 306.355469 C 88.738281 302.375 92.570312 300.421875 97.753906 300.421875 C 106.242188 300.421875 111.953125 306.355469 111.953125 315.671875",
  "M 149.070312 317.023438 C 147.945312 310.9375 143.664062 307.78125 138.332031 307.78125 C 132.320312 307.78125 128.113281 311.3125 126.988281 317.023438 Z M 127.0625 323.636719 C 128.261719 329.34375 132.84375 332.5 139.082031 332.5 C 143.363281 332.5 146.59375 330.921875 148.773438 327.839844 L 155.457031 331.746094 C 151.777344 337.230469 146.292969 339.933594 139.007812 339.933594 C 132.921875 339.933594 127.960938 338.058594 124.207031 334.378906 C 120.449219 330.621094 118.570312 325.886719 118.570312 320.179688 C 118.570312 314.542969 120.449219 309.886719 124.132812 306.128906 C 127.8125 302.300781 132.621094 300.421875 138.40625 300.421875 C 143.890625 300.421875 148.398438 302.375 151.925781 306.207031 C 155.457031 310.035156 157.261719 314.769531 157.261719 320.253906 C 157.261719 321.082031 157.183594 322.207031 156.960938 323.636719 Z M 127.0625 323.636719",
];
function AMLogo({ s, light }) {
  return (
    <svg width={s} height={s} viewBox="0 0 589.606 595.276" style={{ display: "block", flexShrink: 0 }}>
      {!light ? <rect x="0" y="0" width="589.606" height="595.276" rx="90" fill={BC.purple} /> : null}
      {AM_RINGS.map((r, i) => <path key={"r" + i} d={r.d} fill={light ? "#FFFFFF" : r.c} />)}
      {AM_WORDMARK.map((d, i) => <path key={"w" + i} d={d} fill="#FFFFFF" />)}
    </svg>
  );
}

// ---------------- FRAMEWORK DATA ----------------
const RATINGS = ["E", "D", "Es"];
const RLABEL = { E: "Emerging", D: "Developing", Es: "Established" };
const FLAGS = [["pp", "PP", "#944FC9", "#FFFFFF"], ["send", "SEND", "#2554C7", "#FFFFFF"], ["eal", "EAL", "#E06B22", "#FFFFFF"]];
const GROUPS = [
  { label: "All pupils", filter: () => true, bg: BC.ink, fg: "#FFFFFF" },
  { label: "Boys", filter: p => p.g === "B", bg: "#47ABFB", fg: BC.ink },
  { label: "Girls", filter: p => p.g === "G", bg: "#FF6293", fg: BC.ink },
  { label: "PP", filter: p => p.pp, bg: "#944FC9", fg: "#FFFFFF" },
  { label: "SEND", filter: p => p.send, bg: "#2554C7", fg: "#FFFFFF" },
  { label: "EAL", filter: p => p.eal, bg: "#E06B22", fg: "#FFFFFF" },
];

// 6 terms
const TERMS = [
  { n: 1, label: "Term 1", sub: "Autumn 1" },
  { n: 2, label: "Term 2", sub: "Autumn 2" },
  { n: 3, label: "Term 3", sub: "Spring 1" },
  { n: 4, label: "Term 4", sub: "Spring 2" },
  { n: 5, label: "Term 5", sub: "Summer 1" },
  { n: 6, label: "Term 6", sub: "Summer 2" },
];
// Default paired-pillar rotation alongside Physical (which runs every term)
const DEFAULT_ROTATION = ["Social Me", "Mental Me", "Personal Me", "Social Me", "Mental Me", "Personal Me"];
const buildMap = () => TERMS.map((t, i) => ({ term: t.n, paired: DEFAULT_ROTATION[i] }));
const CURRENT_TERM = 3; // demo "today" sits in Spring 1

const FRAMEWORK = {
  "Physical Me": {
    hex: "#2554C7", hexL: "#47ABFB",
    stages: {
      1: { focus: "Move & Try", statements: ["I can move safely in different ways", "I can balance and land with control", "I can find space when playing", "I can use equipment safely"] },
      2: { focus: "Practise & Improve", statements: ["I can run, jump and change direction with control", "I can throw, catch and strike with improving accuracy", "I can choose the right skill for a task", "I can keep trying when activities feel hard"] },
      3: { focus: "Apply Skills", statements: ["I can use skills successfully in games", "I can use simple tactics to help my team", "I can perform movements with good technique", "I can work hard throughout activities"] },
      4: { focus: "Adapt & Perform", statements: ["I can adapt skills for different games and challenges", "I can combine skills confidently under pressure", "I can prepare myself to perform well", "I can help others improve their skills"] },
    },
  },
  "Mental Me": {
    hex: "#189E8A", hexL: "#47FBDF",
    stages: {
      1: { focus: "Listen & Try", statements: ["I can listen carefully to instructions", "I can remember simple rules", "I can try again when something is difficult", "I can talk about what went well"] },
      2: { focus: "Think & Improve", statements: ["I can set myself a simple goal", "I can explain what I need to improve", "I can choose a suitable level of challenge", "I can stay focused during activities"] },
      3: { focus: "Reflect & Adapt", statements: ["I can review my progress towards a goal", "I can change my approach when needed", "I can explain how activity helps my learning and mood", "I can stay focused when challenged"] },
      4: { focus: "Plan & Lead Thinking", statements: ["I can plan strategies to improve performance", "I can evaluate what worked and why", "I can help others stay focused", "I can take responsibility for my progress"] },
    },
  },
  "Social Me": {
    hex: "#E3225C", hexL: "#FF6293",
    stages: {
      1: { focus: "Plays Fairly & Kind", statements: ["I can share equipment and take turns fairly", "I can use kind words to support others", "I can follow rules and be honest in games", "I can ask an adult for help when something feels unsafe"] },
      2: { focus: "Includes Others", statements: ["I can invite others to join in", "I can give kind and helpful feedback", "I can speak up kindly to help keep others safe", "I can show respect when I win or lose"] },
      3: { focus: "Contributes to Team", statements: ["I can play my role to support the team", "I can help solve disagreements fairly", "I can take responsibility for my actions in a group", "I can encourage others and keep a positive attitude"] },
      4: { focus: "Leads & Supports", statements: ["I can lead a group or activity to support others", "I can give and receive feedback to help improve performance", "I can make sure everyone is included, involved and safe", "I can motivate others and keep the group focused"] },
    },
  },
  "Personal Me": {
    hex: "#944FC9", hexL: "#C77EFF",
    stages: {
      1: { focus: "Manage Myself", statements: ["I can follow routines and expectations", "I can join in with a positive attitude", "I can talk about how I feel", "I can try my best"] },
      2: { focus: "Take Responsibility", statements: ["I can set myself a personal goal", "I can use a strategy to stay calm when frustrated", "I can look after my kit and equipment", "I can choose to stay positive"] },
      3: { focus: "Know Myself", statements: ["I can talk about my strengths and next steps", "I can prepare myself for learning", "I can keep going when things are difficult", "I can use what I learn in PE in other situations"] },
      4: { focus: "Grow & Apply", statements: ["I can work towards a longer-term goal", "I can take responsibility for being ready to learn", "I can use resilience and teamwork across school life", "I can describe the person I want to become"] },
    },
  },
};
const PILLARS = Object.keys(FRAMEWORK);
const PAIRED_PILLARS = ["Social Me", "Mental Me", "Personal Me"];
// Headings used to segment the AI-drafted narrative by development area.
const PILLAR_OVERVIEW = {
  "Physical Me": "Physical development overview",
  "Mental Me": "Mental development overview",
  "Social Me": "Social development overview",
  "Personal Me": "Personal development overview",
};

// ---------------- NARRATIVE FORMS (per-statement, grammar-correct) ----------------
// Hand-authored bare-infinitive ("base", used after "can") and gerund ("ger", used
// after "is working towards" / "next target(s) to develop") forms for every
// statement, keyed by the exact "I can..." text. Tokens resolved per pupil gender
// by resolveTemplate(): {poss}=her/his/their, {refl}=herself/himself/themselves,
// {sv:word}=word conjugated to 3rd-person-singular for a named pupil, bare for a
// class ("N"). This avoids the grammar bugs that come from just stripping "I can ".
const NARRATIVE_FORMS = {
  // Physical Me
  "I can move safely in different ways": { base: "move safely in different ways", ger: "moving safely in different ways" },
  "I can balance and land with control": { base: "balance and land with control", ger: "balancing and landing with control" },
  "I can find space when playing": { base: "find space when playing", ger: "finding space when playing" },
  "I can use equipment safely": { base: "use equipment safely", ger: "using equipment safely" },
  "I can run, jump and change direction with control": { base: "run, jump and change direction with control", ger: "running, jumping and changing direction with control" },
  "I can throw, catch and strike with improving accuracy": { base: "throw, catch and strike with improving accuracy", ger: "throwing, catching and striking with improving accuracy" },
  "I can choose the right skill for a task": { base: "choose the right skill for a task", ger: "choosing the right skill for a task" },
  "I can keep trying when activities feel hard": { base: "keep trying when activities feel hard", ger: "keeping trying when activities feel hard" },
  "I can use skills successfully in games": { base: "use skills successfully in games", ger: "using skills successfully in games" },
  "I can use simple tactics to help my team": { base: "use simple tactics to help {poss} team", ger: "using simple tactics to help {poss} team" },
  "I can perform movements with good technique": { base: "perform movements with good technique", ger: "performing movements with good technique" },
  "I can work hard throughout activities": { base: "work hard throughout activities", ger: "working hard throughout activities" },
  "I can adapt skills for different games and challenges": { base: "adapt skills for different games and challenges", ger: "adapting skills for different games and challenges" },
  "I can combine skills confidently under pressure": { base: "combine skills confidently under pressure", ger: "combining skills confidently under pressure" },
  "I can prepare myself to perform well": { base: "prepare {refl} to perform well", ger: "preparing {refl} to perform well" },
  "I can help others improve their skills": { base: "help others improve their skills", ger: "helping others improve their skills" },
  // Mental Me
  "I can listen carefully to instructions": { base: "listen carefully to instructions", ger: "listening carefully to instructions" },
  "I can remember simple rules": { base: "remember simple rules", ger: "remembering simple rules" },
  "I can try again when something is difficult": { base: "try again when something is difficult", ger: "trying again when something is difficult" },
  "I can talk about what went well": { base: "talk about what went well", ger: "talking about what went well" },
  "I can set myself a simple goal": { base: "set {refl} a simple goal", ger: "setting {refl} a simple goal" },
  "I can explain what I need to improve": { base: "explain what {subj} {sv:need} to improve", ger: "explaining what {subj} {sv:need} to improve" },
  "I can choose a suitable level of challenge": { base: "choose a suitable level of challenge", ger: "choosing a suitable level of challenge" },
  "I can stay focused during activities": { base: "stay focused during activities", ger: "staying focused during activities" },
  "I can review my progress towards a goal": { base: "review {poss} progress towards a goal", ger: "reviewing {poss} progress towards a goal" },
  "I can change my approach when needed": { base: "change {poss} approach when needed", ger: "changing {poss} approach when needed" },
  "I can explain how activity helps my learning and mood": { base: "explain how activity helps {poss} learning and mood", ger: "explaining how activity helps {poss} learning and mood" },
  "I can stay focused when challenged": { base: "stay focused when challenged", ger: "staying focused when challenged" },
  "I can plan strategies to improve performance": { base: "plan strategies to improve performance", ger: "planning strategies to improve performance" },
  "I can evaluate what worked and why": { base: "evaluate what worked and why", ger: "evaluating what worked and why" },
  "I can help others stay focused": { base: "help others stay focused", ger: "helping others stay focused" },
  "I can take responsibility for my progress": { base: "take responsibility for {poss} progress", ger: "taking responsibility for {poss} progress" },
  // Social Me
  "I can share equipment and take turns fairly": { base: "share equipment and take turns fairly", ger: "sharing equipment and taking turns fairly" },
  "I can use kind words to support others": { base: "use kind words to support others", ger: "using kind words to support others" },
  "I can follow rules and be honest in games": { base: "follow rules and be honest in games", ger: "following rules and being honest in games" },
  "I can ask an adult for help when something feels unsafe": { base: "ask an adult for help when something feels unsafe", ger: "asking an adult for help when something feels unsafe" },
  "I can invite others to join in": { base: "invite others to join in", ger: "inviting others to join in" },
  "I can give kind and helpful feedback": { base: "give kind and helpful feedback", ger: "giving kind and helpful feedback" },
  "I can speak up kindly to help keep others safe": { base: "speak up kindly to help keep others safe", ger: "speaking up kindly to help keep others safe" },
  "I can show respect when I win or lose": { base: "show respect when {subj} {sv:win} or {sv:lose}", ger: "showing respect when {subj} {sv:win} or {sv:lose}" },
  "I can play my role to support the team": { base: "play {poss} role to support the team", ger: "playing {poss} role to support the team" },
  "I can help solve disagreements fairly": { base: "help solve disagreements fairly", ger: "helping solve disagreements fairly" },
  "I can take responsibility for my actions in a group": { base: "take responsibility for {poss} actions in a group", ger: "taking responsibility for {poss} actions in a group" },
  "I can encourage others and keep a positive attitude": { base: "encourage others, and {sv:have} a positive attitude", ger: "encouraging others, and having a positive attitude" },
  "I can lead a group or activity to support others": { base: "lead a group or activity to support others", ger: "leading a group or activity to support others" },
  "I can give and receive feedback to help improve performance": { base: "give and receive feedback to help improve performance", ger: "giving and receiving feedback to help improve performance" },
  "I can make sure everyone is included, involved and safe": { base: "make sure everyone is included, involved and safe", ger: "making sure everyone is included, involved and safe" },
  "I can motivate others and keep the group focused": { base: "motivate others, and {sv:keep} the group focused", ger: "motivating others, and keeping the group focused" },
  // Personal Me
  "I can follow routines and expectations": { base: "follow routines and expectations", ger: "following routines and expectations" },
  "I can join in with a positive attitude": { base: "join in with a positive attitude", ger: "joining in with a positive attitude" },
  "I can talk about how I feel": { base: "talk about how {subj} {sv:feel}", ger: "talking about how {subj} {sv:feel}" },
  "I can try my best": { base: "try {poss} best", ger: "trying {poss} best" },
  "I can set myself a personal goal": { base: "set {refl} a personal goal", ger: "setting {refl} a personal goal" },
  "I can use a strategy to stay calm when frustrated": { base: "use a strategy to stay calm when frustrated", ger: "using a strategy to stay calm when frustrated" },
  "I can look after my kit and equipment": { base: "look after {poss} kit and equipment", ger: "looking after {poss} kit and equipment" },
  "I can choose to stay positive": { base: "choose to stay positive", ger: "choosing to stay positive" },
  "I can talk about my strengths and next steps": { base: "talk about {poss} strengths and next steps", ger: "talking about {poss} strengths and next steps" },
  "I can prepare myself for learning": { base: "prepare {refl} for learning", ger: "preparing {refl} for learning" },
  "I can keep going when things are difficult": { base: "keep going when things are difficult", ger: "keeping going when things are difficult" },
  "I can use what I learn in PE in other situations": { base: "use what {subj} {sv:learn} in PE in other situations", ger: "using what {subj} {sv:learn} in PE in other situations" },
  "I can work towards a longer-term goal": { base: "work towards a longer-term goal", ger: "working towards a longer-term goal" },
  "I can take responsibility for being ready to learn": { base: "take responsibility for being ready to learn", ger: "taking responsibility for being ready to learn" },
  "I can use resilience and teamwork across school life": { base: "use resilience and teamwork across school life", ger: "using resilience and teamwork across school life" },
  "I can describe the person I want to become": { base: "describe the person {subj} {sv:want} to become", ger: "describing the person {subj} {sv:want} to become" },
};
// Irregular 3rd-person-singular verb forms (regular verbs just add "s").
const SV_IRREGULAR = { have: "has" };
function conjugate(word, singular) {
  if (!singular) return word;
  return SV_IRREGULAR[word] || (word + "s");
}
// Resolves {poss}/{refl}/{sv:word} tokens for a pupil ("G"/"B") or a whole class ("N").
function resolveTemplate(tpl, cat) {
  const subj = cat === "G" ? "she" : cat === "B" ? "he" : "they";
  const poss = cat === "G" ? "her" : cat === "B" ? "his" : "their";
  const refl = cat === "G" ? "herself" : cat === "B" ? "himself" : "themselves";
  const singular = cat !== "N";
  return tpl
    .replace(/\{subj\}/g, subj)
    .replace(/\{poss\}/g, poss)
    .replace(/\{refl\}/g, refl)
    .replace(/\{sv:(\w+)\}/g, (m, w) => conjugate(w, singular));
}

// ---------------- DEMO DATA (fictional) ----------------
const mk = (n, i, f) => {
  const ff = f || {};
  const g = ff.g === "G" ? "G" : ff.g === "B" ? "B" : null;
  return { name: n, init: i, g, pp: !!ff.pp, send: !!ff.send, eal: !!ff.eal };
};
const MAPLE_PUPILS = [
  mk("Amelia K.", "AK", { g: "G", pp: 1 }), mk("Ben T.", "BT", { g: "B", send: 1 }), mk("Chloe M.", "CM", { g: "G", pp: 1, eal: 1 }), mk("Dev P.", "DP", { g: "B" }),
  mk("Eva R.", "ER", { g: "G", pp: 1 }), mk("Finn O.", "FO", { g: "B", send: 1 }), mk("Grace L.", "GL", { g: "G" }), mk("Hassan A.", "HA", { g: "B", eal: 1 }),
  mk("Isla W.", "IW", { g: "G", pp: 1 }), mk("Jack D.", "JD", { g: "B" }), mk("Keira S.", "KS", { g: "G", eal: 1 }), mk("Leo N.", "LN", { g: "B", send: 1 }),
];
const WILLOW_PUPILS = [
  mk("Aaron B.", "AB", { g: "B", pp: 1 }), mk("Bella F.", "BF", { g: "G" }), mk("Caleb H.", "CH", { g: "B", send: 1 }), mk("Daisy J.", "DJ", { g: "G", eal: 1 }),
  mk("Ethan K.", "EK", { g: "B" }), mk("Freya M.", "FM", { g: "G", pp: 1 }), mk("George P.", "GP", { g: "B" }), mk("Holly R.", "HR", { g: "G", pp: 1, send: 1 }),
  mk("Ivan S.", "IS", { g: "B", eal: 1 }), mk("Jasmine T.", "JT", { g: "G" }), mk("Kian V.", "KV", { g: "B" }), mk("Lily Z.", "LZ", { g: "G", pp: 1 }),
];

const withMap = c => ({ ...c, map: c.map || buildMap() });
const INIT_SCHOOLS = [
  {
    id: "stm", name: "St Mary's CE Primary", level: 1, colour: "#2554C7", logo: null,
    classes: [
      { id: "robins", name: "Robins", year: "Year 1", stage: 2, size: 30, pupils: [] },
      { id: "wrens", name: "Wrens", year: "Reception", stage: 1, size: 28, pupils: [] },
    ].map(withMap),
  },
  {
    id: "riv", name: "Riverside Juniors", level: 2, colour: "#189E8A", logo: null,
    classes: [
      { id: "badger", name: "Badger", year: "Year 5", stage: 3, size: 28, pupils: [] },
      { id: "fox", name: "Fox", year: "Year 3", stage: 3, size: 27, pupils: [] },
    ].map(withMap),
  },
  {
    id: "oak", name: "Oakfield Primary", level: 3, colour: "#E3225C", logo: null,
    classes: [
      { id: "maple", name: "4 Maple", year: "Year 4", stage: 3, size: 12, pupils: MAPLE_PUPILS },
      { id: "willow", name: "6 Willow", year: "Year 6", stage: 4, size: 12, pupils: WILLOW_PUPILS },
    ].map(withMap),
  },
];

const USERS = [
  { name: "Will", role: "admin" },
  { name: "Abbey", role: "educator" },
  { name: "Jordan", role: "educator" },
  { name: "Sam", role: "educator" },
  { name: "Priya", role: "educator" },
];

const shiftUp = { E: "D", D: "Es", Es: "Es" };
function seedL3(pupils) {
  const base = {}, re = {};
  pupils.forEach((p, i) => {
    base[p.init] = {}; re[p.init] = {};
    for (let j = 0; j < 4; j++) {
      const r = i >= 10 ? (j % 2 ? "Es" : "D") : i % 4 === 0 ? "D" : j < 2 ? "E" : "D";
      base[p.init][j] = r;
      re[p.init][j] = i % 5 === 4 ? r : shiftUp[r];
    }
  });
  return { base, re };
}
const mapleSeed = seedL3(MAPLE_PUPILS);

// records now carry term
const SEED_ASSESSMENTS = [
  {
    id: 1, schoolId: "riv", classId: "badger", pillar: "Personal Me", stage: 3, window: "base", block: 6, term: 1, date: "2026-04-08", educator: "Abbey", level: 2,
    data: {
      statements: [
        { many: "E", initials: { E: [], D: ["MJ", "GH", "WE"], Es: [] } },
        { many: null, initials: { E: ["TD", "WE", "PO"], D: [], Es: [] } },
        { many: "D", initials: { E: [], D: [], Es: ["SD", "GB", "WE"] } },
        { many: null, initials: { E: [], D: [], Es: [] } },
      ],
      pre: [], post: [],
    },
  },
  {
    id: 2, schoolId: "riv", classId: "badger", pillar: "Personal Me", stage: 3, window: "check", block: 6, term: 1, date: "2026-05-15", educator: "Abbey", level: 2,
    data: {
      statements: [
        { many: "D", initials: { E: [], D: [], Es: [] } },
        { many: "E", initials: { E: [], D: ["TD", "WE", "PO"], Es: [] } },
        { many: "D", initials: { E: [], D: [], Es: [] } },
        { many: "E", initials: { E: [], D: ["HD", "WE", "PO"], Es: [] } },
      ],
      pre: [], post: ["SD"],
    },
  },
  {
    id: 3, schoolId: "stm", classId: "robins", pillar: "Physical Me", stage: 2, window: "base", block: 12, term: 1, date: "2026-04-20", educator: "Jordan", level: 1,
    data: { statements: [{ E: 14, D: 12, Es: 4 }, { E: 16, D: 11, Es: 3 }, { E: 10, D: 15, Es: 5 }, { E: 12, D: 14, Es: 4 }], preCount: 2, postCount: 0 },
  },
  {
    id: 4, schoolId: "stm", classId: "robins", pillar: "Physical Me", stage: 2, window: "check", block: 12, term: 1, date: "2026-06-08", educator: "Jordan", level: 1,
    data: { statements: [{ E: 5, D: 15, Es: 10 }, { E: 6, D: 16, Es: 8 }, { E: 3, D: 14, Es: 13 }, { E: 4, D: 16, Es: 10 }], preCount: 0, postCount: 3 },
  },
  {
    id: 5, schoolId: "oak", classId: "maple", pillar: "Social Me", stage: 3, window: "base", block: 6, term: 1, date: "2026-04-22", educator: "Sam", level: 3,
    data: { pupils: mapleSeed.base, flags: {} },
  },
  {
    id: 6, schoolId: "oak", classId: "maple", pillar: "Social Me", stage: 3, window: "check", block: 6, term: 1, date: "2026-06-05", educator: "Sam", level: 3,
    data: { pupils: mapleSeed.re, flags: { LN: "above" } },
  },
  {
    id: 7, schoolId: "riv", classId: "fox", pillar: "Mental Me", stage: 3, window: "base", block: 12, term: 2, date: "2026-04-21", educator: "Priya", level: 2,
    data: {
      statements: [
        { many: "E", initials: { E: [], D: ["AB", "CD"], Es: [] } },
        { many: "E", initials: { E: [], D: [], Es: [] } },
        { many: "D", initials: { E: ["GH", "JK"], D: [], Es: [] } },
        { many: "E", initials: { E: [], D: ["LM"], Es: [] } },
      ],
      pre: [], post: [],
    },
  },
];

// ---------------- HELPERS ----------------
const WLABEL = { base: "Baseline", check: "Progress Check" };
const termLabel = n => { const t = TERMS.find(x => x.n === n); return t ? t.label + " (" + t.sub + ")" : "Term " + n; };
const getR = (rec, init, j) => {
  const g = rec.data.pupils[init];
  return g ? g[j] : undefined;
};
function dist(rec, j, cls) {
  const out = { E: 0, D: 0, Es: 0 };
  if (!rec) return out;
  if (rec.level === 1) {
    const s = rec.data.statements[j];
    out.E = s.E; out.D = s.D; out.Es = s.Es;
  } else if (rec.level === 2) {
    const s = rec.data.statements[j];
    let named = 0;
    RATINGS.forEach(r => { out[r] = s.initials[r].length; named += s.initials[r].length; });
    if (s.many) out[s.many] += Math.max(cls.size - named, 0);
  } else {
    Object.values(rec.data.pupils).forEach(g => { if (g[j]) out[g[j]]++; });
  }
  return out;
}
const total = d => d.E + d.D + d.Es;
const pct = (n, t) => (t ? Math.round((n / t) * 100) : 0);
const meanScore = d => (total(d) ? (d.E + 2 * d.D + 3 * d.Es) / total(d) : null);
// Turns a 1-3 rating score into a % position along an Emerging-Developing-Established
// scale (for the at-a-glance progress gauge), and into the matching plain word.
const scalePct = score => Math.max(0, Math.min(100, ((score - 1) / 2) * 100));
const ratingWordAt = score => (score >= 2.5 ? "Established" : score >= 1.5 ? "Developing" : "Emerging");
// Bands the average rating-level movement so "+0.75" reads as good/bad without
// external context. Max possible movement is +2.00 (every statement Emerging -> Established).
const MAX_UPLIFT = 2;
function progressBand(u) {
  if (u === null || u === undefined) return null;
  if (u >= 1) return { label: "Excellent progress", bg: BC.lime, fg: BC.ink, range: "1.00+", def: "children moving an average of 1 or more rating levels" };
  if (u >= 0.5) return { label: "Good progress", bg: "#189E8A", fg: "#FFFFFF", range: "0.50–0.99", def: "children moving an average of between 0.5 and just under 1 rating level" };
  return { label: "Limited progress", bg: "#FFD400", fg: BC.ink, range: "under 0.50", def: "children moving an average of less than 0.5 of a rating level" };
}
const fmtDate = d => { const p = d.split("-"); return p[2] + "/" + p[1] + "/" + p[0].slice(2); };
const today = () => new Date().toISOString().slice(0, 10);
const slug = s => s.toLowerCase().replace(/[^a-z0-9]+/g, "-") + "-" + Math.floor(Math.random() * 999);

function pairUplift(pre, post, cls) {
  let sum = 0, n = 0;
  for (let j = 0; j < 4; j++) {
    const mp = meanScore(dist(pre, j, cls));
    const mq = meanScore(dist(post, j, cls));
    if (mp !== null && mq !== null) { sum += mq - mp; n++; }
  }
  return n ? sum / n : null;
}
function buildPairs(schools, assessments) {
  const out = [];
  schools.forEach(sch => {
    sch.classes.forEach(c => {
      const recs = assessments.filter(a => a.schoolId === sch.id && a.classId === c.id);
      const keys = [...new Set(recs.map(r => r.pillar + "|" + r.stage + "|" + (r.term || 0)))];
      keys.forEach(k => {
        const parts = k.split("|");
        const pillar = parts[0], stage = Number(parts[1]), term = Number(parts[2]);
        const pre = recs.filter(r => r.pillar === pillar && Number(r.stage) === stage && (r.term || 0) === term && r.window === "base").pop();
        const post = recs.filter(r => r.pillar === pillar && Number(r.stage) === stage && (r.term || 0) === term && r.window === "check").pop();
        const uplift = pre && post ? pairUplift(pre, post, c) : null;
        out.push({ school: sch, cls: c, pillar, stage, term, pre, post, uplift, block: (post && post.block) || (pre && pre.block) || null });
      });
    });
  });
  return out;
}
function aggRec(rec, cls) {
  const t = { E: 0, D: 0, Es: 0 };
  for (let j = 0; j < 4; j++) { const d = dist(rec, j, cls); t.E += d.E; t.D += d.D; t.Es += d.Es; }
  return t;
}
function groupAgg(pairsL3, filter) {
  const pre = { E: 0, D: 0, Es: 0 }, post = { E: 0, D: 0, Es: 0 };
  const ids = new Set();
  pairsL3.forEach(p => {
    p.cls.pupils.filter(filter).forEach(pu => {
      ids.add(p.cls.id + "/" + pu.init);
      for (let j = 0; j < 4; j++) {
        const r1 = getR(p.pre, pu.init, j); if (r1) pre[r1]++;
        const r2 = getR(p.post, pu.init, j); if (r2) post[r2]++;
      }
    });
  });
  return { pre, post, n: ids.size };
}
// status of a given class+pillar+term: 'complete' | 'awaiting' | 'none'
function unitStatus(assessments, sch, c, pillar, term) {
  const recs = assessments.filter(a => a.schoolId === sch.id && a.classId === c.id && a.pillar === pillar && (a.term || 0) === term);
  const hasBase = recs.some(r => r.window === "base");
  const hasCheck = recs.some(r => r.window === "check");
  if (hasBase && hasCheck) return "complete";
  if (hasBase) return "awaiting";
  return "none";
}

// ---------------- STAGE PLACEMENT (below / at / above expected stage) ----------------
// Every level already tracks pupils working outside the focal stage: Level 1 as
// simple counts, Level 2 as named initials, Level 3 as per-pupil flags. This reads
// whichever is present so "on track for their stage" can be shown at any level.
function countBelow(rec) {
  if (!rec) return 0;
  if (rec.level === 1) return rec.data.preCount || 0;
  if (rec.level === 2) return (rec.data.pre || []).length;
  return Object.values(rec.data.flags || {}).filter(f => f === "below").length;
}
function countAbove(rec) {
  if (!rec) return 0;
  if (rec.level === 1) return rec.data.postCount || 0;
  if (rec.level === 2) return (rec.data.post || []).length;
  return Object.values(rec.data.flags || {}).filter(f => f === "above").length;
}
function stagePlacement(pairs) {
  const base = { below: 0, at: 0, above: 0, total: 0 };
  const check = { below: 0, at: 0, above: 0, total: 0 };
  pairs.forEach(p => {
    const size = p.cls.size;
    if (p.pre) {
      const below = countBelow(p.pre), above = countAbove(p.pre);
      base.below += below; base.above += above; base.at += Math.max(size - below - above, 0); base.total += size;
    }
    if (p.post) {
      const below = countBelow(p.post), above = countAbove(p.post);
      check.below += below; check.above += above; check.at += Math.max(size - below - above, 0); check.total += size;
    }
  });
  return { base, check };
}

// ---------------- AI-DRAFTED IMPACT NARRATIVE ----------------
// Reads the Progress Check ratings and drafts plain-English impact text, so
// educators aren't left writing this up by hand. Groups each "I can..." statement
// by its rating tier into a fixed three-part structure: confirmed at Established,
// working towards at Developing, next targets at Emerging.
function stripICan(s) {
  return s.replace(/^i can\s+/i, "");
}
// Looks up the grammar-correct form for one statement, resolved for a pupil's
// gender ("G"/"B") or a whole class ("N"). Falls back to a plain strip if a
// statement is ever added without a matching NARRATIVE_FORMS entry.
function narrativeForm(stmt, tier, cat) {
  const forms = NARRATIVE_FORMS[stmt];
  if (!forms) return stripICan(stmt);
  return resolveTemplate(forms[tier], cat);
}
// Oxford comma before the final item, so a trailing statement that's itself a
// compound ("X, and Y") doesn't collide with the list's own "and".
function joinAnd(arr) {
  if (arr.length === 0) return "";
  if (arr.length === 1) return arr[0];
  if (arr.length === 2) return arr[0] + " and " + arr[1];
  return arr.slice(0, -1).join(", ") + ", and " + arr[arr.length - 1];
}
function pupilNarrative(pupil, stage, b, re) {
  if (!re) return null;
  const first = pupil.name.split(" ")[0];
  const cat = pupil.g === "G" ? "G" : pupil.g === "B" ? "B" : "N";
  const byLevel = { Es: [], D: [], E: [] };
  for (let j = 0; j < 4; j++) {
    const r = getR(re, pupil.init, j);
    if (byLevel[r]) byLevel[r].push(stage.statements[j]);
  }
  const sentences = [];
  if (byLevel.Es.length) sentences.push(first + " can confidently and consistently " + joinAnd(byLevel.Es.map(s => narrativeForm(s, "base", cat))) + ".");
  if (byLevel.D.length) sentences.push(first + " is working towards " + joinAnd(byLevel.D.map(s => narrativeForm(s, "ger", cat))) + ".");
  if (byLevel.E.length) sentences.push(first + "\u2019s next target" + (byLevel.E.length > 1 ? "s are" : " is") + " to develop " + joinAnd(byLevel.E.map(s => narrativeForm(s, "ger", cat))) + ".");
  if (!sentences.length) return first + " has not yet been evidenced against this stage\u2019s \u201cI can\u2026\u201d statements.";
  return sentences.join(" ");
}
// Class-level equivalent for Level 1/2 schools, where there's no named pupil data --
// groups each statement by whichever rating the majority of the class sits at.
function classNarrative(stage, cls, rec) {
  if (!rec) return null;
  const byLevel = { Es: [], D: [], E: [] };
  for (let j = 0; j < 4; j++) {
    const d = dist(rec, j, cls);
    const t = total(d);
    if (!t) continue;
    const top = RATINGS.reduce((a, r) => (d[r] > d[a] ? r : a), RATINGS[0]);
    byLevel[top].push({ stmt: stage.statements[j], n: d[top], t });
  }
  const fmtList = (arr, tier) => joinAnd(arr.map(x => narrativeForm(x.stmt, tier, "N") + " (" + x.n + " of " + x.t + " pupils)"));
  const sentences = [];
  if (byLevel.Es.length) sentences.push("Most children can confidently and consistently " + fmtList(byLevel.Es, "base") + ".");
  if (byLevel.D.length) sentences.push("Most children are working towards " + fmtList(byLevel.D, "ger") + ".");
  if (byLevel.E.length) sentences.push("The class\u2019s next targets are to develop " + fmtList(byLevel.E, "ger") + ".");
  return sentences.join(" ") || null;
}

// ---------------- LABELLING & METRIC-EXPLAINER COMPONENTS ----------------
// Small persistent tag so any section -- read on its own, screenshotted, or printed
// out of context -- still says which pillar/stage/term it's about.
function SectionTag({ pillar, stage, term }) {
  return (
    <div className="flex items-center gap-1.5 mb-2 flex-wrap">
      <PillarChip p={pillar} small={true} />
      <span className="text-[10px] font-bold text-slate-400">Stage {stage}{term ? " \u00b7 " + termLabel(term) : ""}</span>
    </div>
  );
}
// One-line plain-English caption placed directly under a metric, so the figure
// above it never has to be interpreted unaided.
function MetricNote({ children }) {
  return <p className="text-[10px] leading-snug mt-1.5" style={{ color: "#8b7f99" }}>{children}</p>;
}
// Visual Emerging -> Developing -> Established scale with Baseline and Progress Check
// markers, so the headline progress figure reads instantly without a caption.
function ProgressGauge({ preScore, postScore, light }) {
  const startPct = scalePct(preScore);
  const endPct = scalePct(postScore);
  const lo = Math.min(startPct, endPct), hi = Math.max(startPct, endPct);
  const track = light ? "rgba(255,255,255,0.2)" : "#ECE6F2";
  const labelColor = light ? BC.lilac : "#94a3b8";
  const fillColor = light ? BC.lime : BC.mid;
  return (
    <div>
      <div className="relative h-2 rounded-full mt-1" style={{ backgroundColor: track }}>
        <div className="absolute top-0 h-2 rounded-full" style={{ left: lo + "%", width: Math.max(hi - lo, 1.5) + "%", backgroundColor: fillColor }}></div>
        <div className="absolute rounded-full border-2" style={{ left: "calc(" + startPct + "% - 5px)", top: -3, width: 10, height: 10, backgroundColor: light ? BC.ink : "#fff", borderColor: fillColor }} title="Baseline"></div>
        <div className="absolute rounded-full" style={{ left: "calc(" + endPct + "% - 5px)", top: -3, width: 10, height: 10, backgroundColor: fillColor }} title="Progress Check"></div>
      </div>
      <div className="flex justify-between text-[9px] font-bold uppercase mt-1.5" style={{ color: labelColor }}>
        <span>Emerging</span><span>Developing</span><span>Established</span>
      </div>
    </div>
  );
}
// Stacked below/at/above-expected-stage bar for one timepoint (Baseline or Progress Check).
function StagePlacementBar({ label, d }) {
  if (!d.total) return (
    <div className="flex items-center gap-2 mb-1.5">
      <span className="w-28 text-[11px] font-bold text-slate-400">{label}</span>
      <span className="text-[11px] text-slate-300">No data yet</span>
    </div>
  );
  const pBelow = pct(d.below, d.total), pAbove = pct(d.above, d.total);
  const pAt = Math.max(100 - pBelow - pAbove, 0);
  return (
    <div className="flex items-center gap-2 mb-1.5">
      <span className="w-28 text-[11px] font-bold text-slate-500">{label}</span>
      <div className="flex-1 h-4 rounded-full overflow-hidden flex" style={{ backgroundColor: "#ECE6F2" }}>
        {pBelow > 0 ? <div className="h-full flex items-center justify-center text-[9px] font-bold" style={{ width: pBelow + "%", backgroundColor: "#FFAC38", color: BC.ink }}>{pBelow}%</div> : null}
        {pAt > 0 ? <div className="h-full flex items-center justify-center text-[9px] font-bold" style={{ width: pAt + "%", backgroundColor: "#189E8A", color: "#fff" }}>{pAt}%</div> : null}
        {pAbove > 0 ? <div className="h-full flex items-center justify-center text-[9px] font-bold" style={{ width: pAbove + "%", backgroundColor: "#47ABFB", color: BC.ink }}>{pAbove}%</div> : null}
      </div>
      <span className="text-[10px] text-slate-400 w-16 text-right">n={d.total}</span>
    </div>
  );
}
// A single rating chip, dimmed for the "before" value -- used in the pupil x
// statement table to show Baseline -> Progress Check at a glance.
function RatingCell({ rb, rr }) {
  const chip = (r, dim) => r ? (
    <span className="inline-flex items-center justify-center rounded text-[9px] font-black w-6 h-5 shrink-0" style={{ backgroundColor: RCOL[r], color: RTXT[r], opacity: dim ? 0.55 : 1 }}>{r}</span>
  ) : <span className="inline-flex items-center justify-center rounded text-[9px] font-bold w-6 h-5 shrink-0 border border-dashed" style={{ borderColor: "#E5E0EB", color: "#cbd5e1" }}>\u2013</span>;
  return (
    <div className="flex items-center gap-0.5 justify-center">
      {chip(rb, true)}
      <ChevronRight size={9} className="text-slate-300 shrink-0" />
      {chip(rr, false)}
    </div>
  );
}
// The reworked per-pupil view: statements across the top, pupils down the side,
// colour-coded Baseline -> Check rating per statement (replaces a single averaged score).
function PupilStatementTable({ pupils, stage, b, re }) {
  return (
    <div>
      <div className="overflow-x-auto -mx-1 px-1">
        <table className="border-collapse" style={{ minWidth: 380, width: "100%" }}>
          <thead>
            <tr>
              <th className="text-left text-[10px] font-bold uppercase px-1 py-1 sticky left-0 bg-white" style={{ color: BC.purple }}>Pupil</th>
              {stage.statements.map((st, j) => (
                <th key={j} className="text-center text-[10px] font-bold px-1 py-1" style={{ color: BC.purple }} title={st}>S{j + 1}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {pupils.map(p => (
              <tr key={p.init} className="border-t" style={{ borderColor: "#F0EBF5" }}>
                <td className="py-1.5 pr-2 sticky left-0 bg-white align-top">
                  <div className="font-medium whitespace-nowrap text-xs" style={{ color: BC.ink }}>{p.name}</div>
                  <FlagChips p={p} small={true} />
                </td>
                {stage.statements.map((st, j) => (
                  <td key={j} className="py-1.5 px-0.5 align-top"><RatingCell rb={b ? getR(b, p.init, j) : undefined} rr={re ? getR(re, p.init, j) : undefined} /></td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <div className="mt-2 space-y-0.5">
        {stage.statements.map((st, j) => <p key={j} className="text-[10px] text-slate-400"><b>S{j + 1}:</b> {st}</p>)}
      </div>
      <p className="text-[10px] text-slate-400 mt-2">Each cell reads Baseline \u2192 Progress Check for that statement \u2014 <span className="font-bold" style={{ color: BC.ink }}>E</span> Emerging, <span className="font-bold" style={{ color: BC.ink }}>D</span> Developing, <span className="font-bold" style={{ color: BC.ink }}>Es</span> Established, "\u2013" not yet evidenced.</p>
    </div>
  );
}

const TEMPLATE_CSV = "Class,Year,Stage,Pupil Name,Initials,Gender,PP,SEND,EAL\n5 Cedar,Year 5,3,Sample Pupil One,SP1,B,Y,,\n5 Cedar,Year 5,3,Sample Pupil Two,SP2,G,,Y,Y\n3 Elm,Year 3,2,Sample Pupil Three,SP3,G,,,\n";
const yesish = v => ["y", "yes", "1", "true", "x"].indexOf(String(v).trim().toLowerCase()) >= 0;
const pick = (row, names) => {
  for (let i = 0; i < names.length; i++) { if (row[names[i]] !== undefined && String(row[names[i]]).trim() !== "") return String(row[names[i]]).trim(); }
  return "";
};
const parseGender = v => {
  const s = String(v).trim().toLowerCase();
  if (["b", "boy", "m", "male"].indexOf(s) >= 0) return "B";
  if (["g", "girl", "f", "female"].indexOf(s) >= 0) return "G";
  return null;
};

// ---------------- SMALL COMPONENTS ----------------
function Bar({ d, size }) {
  const blank = Math.max(size - total(d), 0);
  const full = (total(d) + blank) || 1;
  return (
    <div className="flex h-6 w-full rounded-md overflow-hidden" style={{ backgroundColor: "#ECE6F2" }}>
      {RATINGS.map(r => d[r] > 0 ? (
        <div key={r} className="flex items-center justify-center text-xs font-semibold" style={{ width: ((d[r] / full) * 100) + "%", backgroundColor: RCOL[r], color: RTXT[r] }}>{d[r]}</div>
      ) : null)}
      {blank > 0 ? <div className="flex items-center justify-center text-slate-400 text-xs" style={{ width: ((blank / full) * 100) + "%" }} title="Not yet evidenced">{blank}</div> : null}
    </div>
  );
}
function Counter({ value, onChange, max }) {
  const atMax = max !== undefined && value >= max;
  return (
    <div className="flex items-center gap-1">
      <button onClick={() => onChange(Math.max(0, value - 1))} className="w-9 h-9 rounded-lg flex items-center justify-center active:scale-95" style={{ backgroundColor: "#ECE6F2", color: BC.ink }}><Minus size={16} /></button>
      <div className="w-10 text-center font-bold text-lg tabular-nums" style={{ color: BC.ink }}>{value}</div>
      <button onClick={() => { if (!atMax) onChange(value + 1); }} disabled={atMax} className="w-9 h-9 rounded-lg flex items-center justify-center active:scale-95 hover:opacity-90" style={atMax ? { backgroundColor: "#ECE6F2", color: "#cbd5e1" } : { backgroundColor: BC.ink, color: "#fff" }}><Plus size={16} /></button>
    </div>
  );
}
function InitialChips({ list, onRemove }) {
  return (
    <div className="flex flex-wrap gap-1 mt-1">
      {list.map((i, k) => (
        <span key={k} className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-semibold" style={pill(BC.purple, "#fff")}>
          {i}<button onClick={() => onRemove(k)} className="opacity-70 hover:opacity-100"><X size={12} /></button>
        </span>
      ))}
    </div>
  );
}
function RatingChip({ r }) {
  return <span className="text-xs font-bold px-2 py-1 rounded-md border" style={{ backgroundColor: RCOL[r] + "33", borderColor: RCOL[r], color: BC.ink }}>{RLABEL[r]}</span>;
}
function MiniRating({ r }) {
  return <span className="text-[10px] font-bold px-1.5 py-0.5 rounded border" style={{ backgroundColor: RCOL[r] + "33", borderColor: RCOL[r], color: BC.ink }}>{RLABEL[r]}</span>;
}
function LevelBadge({ level }) {
  const bg = level === 1 ? "#47ABFB" : level === 2 ? "#47FBDF" : "#C77EFF";
  const t = level === 1 ? "Level 1 · Class totals" : level === 2 ? "Level 2 · Class + initials" : "Level 3 · Per pupil";
  return <span className="px-2 py-0.5 rounded-full text-xs font-bold" style={pill(bg, BC.ink)}>{t}</span>;
}
function PillarChip({ p, small }) {
  return <span className={(small ? "text-[11px]" : "text-xs") + " font-bold px-2 py-0.5 rounded-full"} style={pill(FRAMEWORK[p].hexL, BC.ink)}>{p}</span>;
}
function FlagChips({ p, small }) {
  const sz = small ? "text-[9px] px-1" : "text-[10px] px-1.5";
  return (
    <span className="inline-flex gap-1">
      {p.g ? <span className={sz + " py-0.5 rounded font-bold"} style={p.g === "B" ? pill("#47ABFB", BC.ink) : pill("#FF6293", BC.ink)}>{p.g}</span> : null}
      {FLAGS.map(fl => p[fl[0]] ? <span key={fl[0]} className={sz + " py-0.5 rounded font-bold"} style={pill(fl[2], fl[3])}>{fl[1]}</span> : null)}
    </span>
  );
}
function GroupBars({ rows }) {
  return (
    <div className="space-y-2">
      {rows.map(r => (
        <div key={r.label} className="flex items-center gap-3">
          <span className="w-20 text-[11px] font-bold px-2 py-1 rounded text-center" style={pill(r.bg, r.fg)}>{r.label}</span>
          <span className="text-[10px] text-slate-400 w-8">n={r.n}</span>
          <div className="flex-1 h-4 rounded-full overflow-hidden relative" style={{ backgroundColor: "#ECE6F2" }}>
            <div className="absolute inset-y-0 left-0 rounded-full" style={{ width: r.pR + "%", backgroundColor: "#47FBDF" }}></div>
            <div className="absolute inset-y-0 left-0 rounded-full" style={{ width: r.pB + "%", backgroundColor: "#189E8A" }}></div>
          </div>
          <span className="text-xs font-black w-20 text-right tabular-nums" style={{ color: BC.ink }}>{r.pB}% → {r.pR}%</span>
        </div>
      ))}
    </div>
  );
}

// Curriculum map strip: Physical (always) + paired pillar per term
function CurriculumMap({ map, current, onEdit, statusFn }) {
  return (
    <div className="rounded-2xl overflow-hidden border" style={{ borderColor: "#E5E0EB" }}>
      <div className="grid grid-cols-6">
        {map.map(m => {
          const isNow = m.term === current;
          const t = TERMS.find(x => x.n === m.term);
          return (
            <div key={m.term} className="border-r last:border-r-0" style={{ borderColor: "#E5E0EB", backgroundColor: isNow ? BC.lime + "33" : "#fff" }}>
              <div className="text-center py-1" style={{ backgroundColor: isNow ? BC.ink : "#F6F3F8" }}>
                <div className="text-[10px] font-black" style={{ color: isNow ? "#fff" : BC.ink }}>T{m.term}</div>
                <div className="text-[8px]" style={{ color: isNow ? BC.lilac : "#94a3b8" }}>{t.sub}</div>
              </div>
              <div className="p-1 space-y-1">
                {/* Physical always */}
                <PillStack pillar="Physical Me" status={statusFn ? statusFn("Physical Me", m.term) : null} onEdit={onEdit ? () => onEdit(m.term, "Physical Me") : null} />
                <PillStack pillar={m.paired} status={statusFn ? statusFn(m.paired, m.term) : null} onEdit={onEdit ? () => onEdit(m.term, "paired") : null} paired />
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}
function PillStack({ pillar, status, onEdit, paired }) {
  const dot = status === "complete" ? "#189E8A" : status === "awaiting" ? "#E06B22" : null;
  return (
    <button onClick={onEdit || undefined} disabled={!onEdit} className="w-full rounded-md px-1 py-1 text-[9px] font-bold leading-tight relative" style={pill(FRAMEWORK[pillar].hex, "#fff")}>
      {pillar.split(" ")[0]}
      {dot ? <span className="absolute top-0.5 right-0.5 w-1.5 h-1.5 rounded-full" style={{ backgroundColor: dot }}></span> : null}
      {paired && onEdit ? <Pencil size={8} className="absolute bottom-0.5 right-0.5 opacity-70" /> : null}
    </button>
  );
}

// ---------------- MAIN APP ----------------
export default function App() {
  const [screen, setScreen] = useState("login");
  const [user, setUser] = useState(null);
  const [authLoading, setAuthLoading] = useState(true);
  const [authError, setAuthError] = useState("");
  const [loginEmail, setLoginEmail] = useState("");
  const [loginPassword, setLoginPassword] = useState("");
  const [schools, setSchools] = useState([]);
  const [schoolId, setSchoolId] = useState(null);
  const [tab, setTab] = useState("plan");
  const [classId, setClassId] = useState(null);
  const [assessments, setAssessments] = useState([]);
  const [dataLoading, setDataLoading] = useState(false);
  const [cfg, setCfg] = useState(null);
  const [draft, setDraft] = useState(null);
  const [report, setReport] = useState(null);
  const [docScope, setDocScope] = useState(null);
  const [stTab, setStTab] = useState(0);
  const [initInput, setInitInput] = useState({});
  const [saved, setSaved] = useState(false);
  const [adminSchoolId, setAdminSchoolId] = useState(null);
  const [adminClassId, setAdminClassId] = useState(null);
  const [forms, setForms] = useState({});
  const [importPreview, setImportPreview] = useState(null);
  const [scopePick, setScopePick] = useState(null);
  const [showStatements, setShowStatements] = useState(true);
  const [editMap, setEditMap] = useState(null); // {term} for editing paired pillar

  const school = schools.find(s => s.id === schoolId);
  const cls = school ? school.classes.find(c => c.id === classId) : null;
  const isAdmin = user !== null && user.role === "admin";
  const F = (k, d) => (forms[k] === undefined ? (d === undefined ? "" : d) : forms[k]);
  const setF = (k, v) => setForms(f => ({ ...f, [k]: v }));

  // ---------------- SUPABASE: AUTH BOOTSTRAP ----------------
  // On load, and whenever the Supabase auth state changes (sign in / sign out / token
  // refresh), look up the matching row in app_users to get this person's display name
  // and role (educator/admin) - the rest of the app is written around that {name, role}
  // shape, so nothing downstream needs to change.
  useEffect(() => {
    let active = true;
    async function loadProfile(authUser) {
      if (!authUser) {
        if (active) { setUser(null); setAuthLoading(false); }
        return;
      }
      const { data, error } = await supabase
        .from("app_users")
        .select("id,name,role,active")
        .eq("auth_user_id", authUser.id)
        .maybeSingle();
      if (!active) return;
      if (error || !data) {
        setAuthError(error ? error.message : "No account found for this login. Ask an admin to set one up for you.");
        setUser(null);
      } else if (!data.active) {
        setAuthError("This account has been deactivated. Contact an admin.");
        setUser(null);
      } else {
        setAuthError("");
        setUser({ name: data.name, role: data.role, appUserId: data.id });
        setScreen(s => (s === "login" ? "schools" : s));
      }
      setAuthLoading(false);
    }
    supabase.auth.getSession().then(({ data }) => loadProfile(data.session ? data.session.user : null));
    const { data: sub } = supabase.auth.onAuthStateChange((_event, session) => {
      loadProfile(session ? session.user : null);
    });
    return () => { active = false; sub.subscription.unsubscribe(); };
  }, []);

  // ---------------- SUPABASE: DATA LOADING ----------------
  // Pulls schools/classes/pupils and assessments from Supabase and reshapes them into
  // the same nested objects the rest of the app already expects, so the Report/Doc/
  // Entry screens etc. don't need to change at all.
  async function loadAllData() {
    setDataLoading(true);
    const { data: schoolRows } = await supabase.from("schools").select("id,name,level");
    const { data: classRows } = await supabase.from("classes").select("id,school_id,name,year_group,stage,size,curriculum_map");
    const { data: pupilRows } = await supabase.from("pupils").select("id,class_id,init");

    const levelBySchool = {};
    (schoolRows || []).forEach(s => { levelBySchool[s.id] = s.level; });

    const pupilsByClass = {};
    (pupilRows || []).forEach(p => { (pupilsByClass[p.class_id] = pupilsByClass[p.class_id] || []).push({ init: p.init, id: p.id }); });

    const classesBySchool = {};
    (classRows || []).forEach(c => {
      (classesBySchool[c.school_id] = classesBySchool[c.school_id] || []).push({
        id: c.id,
        name: c.name,
        year: c.year_group,
        stage: c.stage,
        size: c.size,
        pupils: pupilsByClass[c.id] || [],
        map: c.curriculum_map && Object.keys(c.curriculum_map).length ? c.curriculum_map : buildMap(),
      });
    });

    setSchools((schoolRows || []).map(s => ({
      id: s.id, name: s.name, level: s.level, classes: classesBySchool[s.id] || [],
    })));

    const { data: assessRows } = await supabase
      .from("assessments")
      .select("id,school_id,class_id,pillar,stage,window_type,block,term,assess_date,educator_id,data,app_users(name)");

    setAssessments((assessRows || []).map(a => ({
      id: a.id,
      schoolId: a.school_id,
      classId: a.class_id,
      pillar: a.pillar,
      stage: a.stage,
      window: a.window_type,
      block: a.block,
      term: a.term,
      date: a.assess_date,
      educator: a.app_users ? a.app_users.name : "",
      level: levelBySchool[a.school_id],
      data: a.data,
    })));
    setDataLoading(false);
  }

  useEffect(() => {
    if (user) loadAllData();
  }, [user]);

  const COLOURS = ["#4A1671", "#189E8A", "#E06B22", "#E3225C", "#2554C7"];
  function addSchool() {
    const name = String(F("sName")).trim(); if (!name) return;
    setSchools(s => [...s, { id: slug(name), name, level: Number(F("sLevel", 1)), colour: COLOURS[s.length % COLOURS.length], logo: null, classes: [] }]);
    setF("sName", ""); setF("sLevel", 1);
  }
  function addClass(sid) {
    const name = String(F("cName")).trim(); if (!name) return;
    const nc = { id: slug(name), name, year: F("cYear", "Year 1"), stage: Number(F("cStage", 1)), size: Number(F("cSize", 30)) || 30, pupils: [], map: buildMap() };
    setSchools(ss => ss.map(s => (s.id !== sid ? s : { ...s, classes: [...s.classes, nc] })));
    setF("cName", "");
  }
  function setPaired(sid, cid, term, pillar) {
    setSchools(ss => ss.map(s => (s.id !== sid ? s : { ...s, classes: s.classes.map(c => (c.id !== cid ? c : { ...c, map: c.map.map(m => (m.term === term ? { ...m, paired: pillar } : m)) })) })));
  }
  function addPupil(sid, cid) {
    const name = String(F("pName")).trim(); if (!name) return;
    const typed = String(F("pInit")).trim().toUpperCase();
    const auto = name.split(" ").map(w => w[0]).join("").slice(0, 3).toUpperCase();
    const np = mk(name, typed || auto, { g: F("pG", null), pp: F("pPP", false), send: F("pSEND", false), eal: F("pEAL", false) });
    setSchools(ss => ss.map(s => (s.id !== sid ? s : { ...s, classes: s.classes.map(c => (c.id !== cid ? c : { ...c, pupils: [...c.pupils, np] })) })));
    setF("pName", ""); setF("pInit", ""); setF("pG", null); setF("pPP", false); setF("pSEND", false); setF("pEAL", false);
  }
  function removePupil(sid, cid, init) {
    setSchools(ss => ss.map(s => (s.id !== sid ? s : { ...s, classes: s.classes.map(c => (c.id !== cid ? c : { ...c, pupils: c.pupils.filter(p => p.init !== init) })) })));
  }
  function setLogo(sid, dataUrl) {
    setSchools(ss => ss.map(s => (s.id !== sid ? s : { ...s, logo: dataUrl })));
  }

  function downloadTemplate() {
    const blob = new Blob([TEMPLATE_CSV], { type: "text/csv" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = "ActiveMe_Pupil_Import_Template.csv";
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }
  async function handleImportFile(e, sid) {
    const file = e.target.files && e.target.files[0];
    e.target.value = "";
    if (!file) return;
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf);
      const raw = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: "" });
      const rows = raw.map(r => {
        const o = {};
        Object.keys(r).forEach(k => { o[k.toLowerCase().trim()] = r[k]; });
        return o;
      });
      const errors = [];
      const classes = {};
      rows.forEach((r, idx) => {
        const cname = pick(r, ["class", "class name"]);
        const pname = pick(r, ["pupil name", "name", "pupil"]);
        if (!cname || !pname) { errors.push("Row " + (idx + 2) + ": missing " + (!cname ? "Class" : "Pupil Name") + " — skipped"); return; }
        const year = pick(r, ["year", "year group"]) || "Year 1";
        let stg = Number(pick(r, ["stage", "focal stage"]));
        if (!(stg >= 1 && stg <= 4)) stg = 1;
        let init = pick(r, ["initials", "init"]).toUpperCase();
        if (!init) init = pname.split(" ").map(w => w[0]).join("").slice(0, 3).toUpperCase();
        if (!classes[cname]) classes[cname] = { name: cname, year, stage: stg, pupils: [] };
        const c = classes[cname];
        if (c.pupils.some(p => p.init === init)) {
          let n = 2;
          while (c.pupils.some(p => p.init === init + n)) n++;
          errors.push("Row " + (idx + 2) + ": duplicate initials '" + init + "' in " + cname + " — saved as " + init + n);
          init = init + n;
        }
        c.pupils.push(mk(pname, init, { g: parseGender(pick(r, ["gender", "sex"])), pp: yesish(pick(r, ["pp", "pupil premium"])), send: yesish(pick(r, ["send", "sen"])), eal: yesish(pick(r, ["eal"])) }));
      });
      const list = Object.values(classes);
      if (list.length === 0) { setImportPreview({ sid, list: [], errors: ["No valid rows found. Check the file uses the template headings: Class, Year, Stage, Pupil Name, Initials, Gender, PP, SEND, EAL."] }); return; }
      setImportPreview({ sid, list, errors });
    } catch (err) {
      setImportPreview({ sid, list: [], errors: ["Could not read the file. Please use the .csv template or an .xlsx file with the same headings."] });
    }
  }
  async function confirmImport() {
    const ip = importPreview;
    const sch = schools.find(s => s.id === ip.sid);
    for (const ic of ip.list) {
      const existing = sch.classes.find(c => c.name.toLowerCase() === ic.name.toLowerCase());
      const classId = existing ? existing.id : slug(ic.name);
      if (!existing) {
        const { error } = await supabase.from("classes").insert({
          id: classId, school_id: ip.sid, name: ic.name, year_group: ic.year, stage: ic.stage,
          size: ic.pupils.length, curriculum_map: buildMap(),
        });
        if (error) { alert("Could not create class '" + ic.name + "': " + error.message); continue; }
      }
      const existingInits = new Set((existing ? existing.pupils : []).map(p => p.init));
      const newPupils = ic.pupils.filter(p => !existingInits.has(p.init));
      if (newPupils.length) {
        const { error } = await supabase.from("pupils").insert(newPupils.map(p => ({ class_id: classId, init: p.init })));
        if (error) alert("Could not add pupils to '" + ic.name + "': " + error.message);
        if (existing) {
          await supabase.from("classes").update({ size: Math.max(existing.size, existing.pupils.length + newPupils.length) }).eq("id", classId);
        }
      }
    }
    await loadAllData();
    setImportPreview(null);
  }

  function newDraft(level) {
    if (level === 1) return { statements: [0, 1, 2, 3].map(() => ({ E: 0, D: 0, Es: 0 })), preCount: 0, postCount: 0 };
    if (level === 2) return { statements: [0, 1, 2, 3].map(() => ({ many: null, initials: { E: [], D: [], Es: [] } })), pre: [], post: [] };
    return { pupils: Object.fromEntries(cls.pupils.map(p => [p.init, {}])), flags: {} };
  }
  async function saveDraft() {
    const { data: inserted, error } = await supabase.from("assessments").insert({
      school_id: schoolId, class_id: classId, pillar: cfg.pillar, stage: cfg.stage,
      window_type: cfg.window, block: cfg.block, term: cfg.term, assess_date: cfg.date,
      educator_id: user.appUserId, data: draft,
    }).select().single();
    if (error) { alert("Could not save this assessment: " + error.message); return; }
    setAssessments(a => [...a, { id: inserted.id, schoolId, classId, pillar: cfg.pillar, stage: cfg.stage, window: cfg.window, block: cfg.block, term: cfg.term, date: cfg.date, educator: user.name, level: school.level, data: draft }]);
    setSaved(true);
    setTimeout(() => { setSaved(false); setTab("progress"); setScreen("school"); }, 900);
  }

  // ---------------- LOGIN ----------------
  if (screen === "login") return (
    <div className="fb min-h-screen flex items-center justify-center p-4" style={{ backgroundColor: BC.ink }}>
      {FONTS}
      <div className="w-full max-w-sm">
        <div className="text-center mb-6">
          <div className="flex justify-center mb-3"><AMLogo s={72} light={true} /></div>
          <h1 className="fd text-white text-2xl font-bold">ActiveMe 360</h1>
          <p className="text-sm mt-1" style={{ color: BC.lilac }}>The ActiveMe Way - Progress Tracker</p>
        </div>
        <div className="bg-white rounded-2xl p-5 shadow-xl">
          {authLoading ? (
            <p className="text-sm text-center py-8" style={{ color: BC.purple }}>Checking your session…</p>
          ) : (
            <>
              {authError ? (
                <div className="flex items-start gap-2 rounded-lg p-2.5 mb-4 border" style={tintBox("#E3225C")}>
                  <ShieldAlert size={18} className="shrink-0 mt-0.5" style={{ color: "#E3225C" }} />
                  <p className="text-xs" style={{ color: BC.ink }}>{authError}</p>
                </div>
              ) : null}
              <button
                onClick={() => { setAuthError(""); supabase.auth.signInWithOAuth({ provider: "azure", options: { scopes: "email" } }); }}
                className="fd w-full py-3 rounded-xl font-bold mb-4 hover:opacity-90"
                style={pill(BC.ink, "#fff")}
              >
                Sign in with Microsoft
              </button>
              <div className="flex items-center gap-2 mb-4">
                <div className="flex-1 h-px" style={{ backgroundColor: "#E5E0EB" }} />
                <span className="text-[10px] uppercase font-bold text-slate-400">or</span>
                <div className="flex-1 h-px" style={{ backgroundColor: "#E5E0EB" }} />
              </div>
              <label className="text-xs font-bold uppercase" style={{ color: BC.purple }}>Email</label>
              <input type="email" value={loginEmail} onChange={e => setLoginEmail(e.target.value)} placeholder="you@activeme360.com" className="w-full mt-1 mb-3 px-3 py-2.5 rounded-lg border border-slate-300 focus:outline-none" />
              <label className="text-xs font-bold uppercase" style={{ color: BC.purple }}>Password</label>
              <div className="relative mt-1 mb-4">
                <Lock size={16} className="absolute left-3 top-3 text-slate-400" />
                <input type="password" value={loginPassword} onChange={e => setLoginPassword(e.target.value)} placeholder="Password" className="w-full pl-9 pr-3 py-2.5 rounded-lg border border-slate-300 focus:outline-none" onKeyDown={e => { if (e.key === "Enter" && loginEmail && loginPassword) e.currentTarget.form?.requestSubmit?.(); }} />
              </div>
              <button
                disabled={!loginEmail || !loginPassword}
                onClick={async () => {
                  setAuthError("");
                  const { error } = await supabase.auth.signInWithPassword({ email: loginEmail, password: loginPassword });
                  if (error) setAuthError(error.message);
                }}
                className="fd w-full py-3 rounded-xl font-bold disabled:opacity-30 hover:opacity-90"
                style={pill(BC.lime, BC.ink)}
              >
                Sign in
              </button>
            </>
          )}
        </div>
      </div>
    </div>
  );

  const Header = ({ title, back, sub }) => (
    <div className="sticky top-0 z-10 text-white px-4 py-3 flex items-center gap-3 shadow-md" style={{ backgroundColor: BC.ink }}>
      {back ? <button onClick={back} className="p-1.5 rounded-lg hover:bg-white/10"><ChevronLeft size={20} /></button> : null}
      <div className="flex-1 min-w-0">
        <h1 className="fd font-bold truncate text-sm">{title}</h1>
        {sub ? <p className="text-xs truncate" style={{ color: BC.lilac }}>{sub}</p> : null}
      </div>
      <div className="text-xs px-2.5 py-1 rounded-full font-semibold" style={isAdmin ? pill(BC.lime, BC.ink) : pill(BC.purple, "#fff")}>{user.name}{isAdmin ? " · Admin" : ""}</div>
    </div>
  );

  // ---------------- SCHOOLS LIST ----------------
  if (screen === "schools") return (
    <div className="fb min-h-screen" style={{ backgroundColor: BC.bg }}>
      {FONTS}
      <Header title="My Schools" sub="Select a school to plan, log or review" back={() => setScreen("login")} />
      <div className="p-4 max-w-lg mx-auto space-y-3">
        {isAdmin ? (
          <div className="grid grid-cols-2 gap-3">
            <button onClick={() => { setAdminSchoolId(null); setAdminClassId(null); setScreen("admin"); }} className="text-white rounded-2xl p-4 shadow-sm hover:opacity-95 text-left" style={{ backgroundColor: BC.ink }}>
              <div className="w-10 h-10 rounded-xl flex items-center justify-center mb-2" style={{ backgroundColor: BC.mid }}><Settings size={20} /></div>
              <div className="fd font-bold text-sm">Admin Console</div>
              <div className="text-[11px]" style={{ color: BC.lilac }}>Schools, classes, curriculum maps, imports</div>
            </button>
            <button onClick={() => setScreen("company")} className="text-white rounded-2xl p-4 shadow-sm hover:opacity-95 text-left" style={{ backgroundColor: BC.ink }}>
              <div className="w-10 h-10 rounded-xl flex items-center justify-center mb-2" style={pill(BC.lime, BC.ink)}><Briefcase size={20} /></div>
              <div className="fd font-bold text-sm">Company Dashboard</div>
              <div className="text-[11px]" style={{ color: BC.lilac }}>Whole-company impact & staff</div>
            </button>
          </div>
        ) : null}
        {schools.map(s => (
          <button key={s.id} onClick={() => { setSchoolId(s.id); setClassId(null); setTab("plan"); setScopePick(null); setScreen("school"); }} className="w-full bg-white rounded-2xl p-4 shadow-sm hover:shadow-md text-left flex items-center gap-4 transition-shadow">
            <div className="w-12 h-12 rounded-xl text-white flex items-center justify-center font-black text-lg shrink-0" style={{ backgroundColor: s.colour }}>{s.name[0]}</div>
            <div className="flex-1">
              <div className="font-bold" style={{ color: BC.ink }}>{s.name}</div>
              <div className="text-xs text-slate-500">{s.classes.length} classes</div>
              <div className="mt-1"><LevelBadge level={s.level} /></div>
            </div>
            <ChevronRight className="text-slate-300" />
          </button>
        ))}
      </div>
    </div>
  );

  // ---------------- COMPANY DASHBOARD ----------------
  if (screen === "company") {
    const allPairs = buildPairs(schools, assessments);
    const complete = allPairs.filter(p => p.uplift !== null);
    const classesAssessed = [...new Set(allPairs.map(p => p.school.id + "/" + p.cls.id))];
    const pupilsReached = classesAssessed.reduce((n, key) => {
      const parts = key.split("/");
      const sch = schools.find(s => s.id === parts[0]);
      const c = sch ? sch.classes.find(x => x.id === parts[1]) : null;
      return n + (c ? c.size : 0);
    }, 0);
    const avgUplift = complete.length ? complete.reduce((a, p) => a + p.uplift, 0) / complete.length : 0;
    const byPillar = PILLARS.map(pl => {
      const ps = complete.filter(p => p.pillar === pl);
      return { pillar: pl, n: ps.length, uplift: ps.length ? ps.reduce((a, p) => a + p.uplift, 0) / ps.length : null };
    }).filter(x => x.uplift !== null);
    const staffMap = {};
    complete.forEach(p => {
      const e = p.post.educator;
      if (!staffMap[e]) staffMap[e] = { name: e, pairs: 0, sum: 0 };
      staffMap[e].pairs++; staffMap[e].sum += p.uplift;
    });
    const staff = Object.values(staffMap).map(s => ({ ...s, avg: s.sum / s.pairs })).sort((a, b) => b.avg - a.avg);
    const maxU = Math.max(0.01, ...staff.map(s => s.avg), ...byPillar.map(b => b.uplift));

    const Card = ({ label, children, accent }) => (
      <div className="bg-white rounded-2xl p-4 shadow-sm">
        <div className="text-[10px] font-bold uppercase" style={{ color: accent || "#94a3b8" }}>{label}</div>
        {children}
      </div>
    );

    return (
      <div className="fb min-h-screen pb-10" style={{ backgroundColor: BC.bg }}>
        {FONTS}
        <Header title="Company Dashboard" sub="ActiveMe 360 · Whole-company impact (admin only)" back={() => setScreen("schools")} />
        <div className="max-w-lg mx-auto p-3 space-y-3">
          <div className="grid grid-cols-2 gap-2">
            <Card label="Partner schools"><div className="fd text-3xl font-bold" style={{ color: BC.ink }}>{schools.length}</div></Card>
            <Card label="Pupils reached"><div className="fd text-3xl font-bold" style={{ color: BC.ink }}>{pupilsReached}</div></Card>
            <Card label="Completed units (baseline + check)"><div className="fd text-3xl font-bold" style={{ color: BC.ink }}>{complete.length}<span className="text-base text-slate-400"> / {allPairs.length}</span></div></Card>
            <Card label="Avg progress" accent="#189E8A"><div className="fd text-3xl font-bold" style={{ color: "#189E8A" }}>+{avgUplift.toFixed(2)}</div><div className="text-[10px] text-slate-400">rating levels per unit</div></Card>
          </div>

          <div className="bg-white rounded-2xl p-4 shadow-sm">
            <p className="fd font-bold text-sm mb-3" style={{ color: BC.ink }}>Impact by pillar</p>
            <div className="space-y-2">
              {byPillar.map(b => (
                <div key={b.pillar} className="flex items-center gap-2">
                  <span className="w-24 text-[11px] font-bold px-2 py-1 rounded text-center" style={pill(FRAMEWORK[b.pillar].hexL, BC.ink)}>{b.pillar}</span>
                  <span className="text-[10px] text-slate-400 w-8">n={b.n}</span>
                  <div className="flex-1 h-3.5 rounded-full overflow-hidden" style={{ backgroundColor: "#ECE6F2" }}>
                    <div className="h-full rounded-full" style={{ width: Math.max((b.uplift / maxU) * 100, 3) + "%", backgroundColor: FRAMEWORK[b.pillar].hex }}></div>
                  </div>
                  <span className="text-xs font-black w-12 text-right tabular-nums" style={{ color: BC.ink }}>+{b.uplift.toFixed(2)}</span>
                </div>
              ))}
            </div>
            <p className="text-[11px] text-slate-400 mt-2">Average movement in rating levels (Emerging=1, Developing=2, Established=3) across statements checked at both Baseline and Progress Check.</p>
          </div>

          <div className="bg-white rounded-2xl p-4 shadow-sm">
            <p className="fd font-bold text-sm mb-1" style={{ color: BC.ink }}>Staff impact</p>
            <p className="text-[11px] text-slate-400 mb-3">Ranked by average pupil progress across their completed units. Use alongside context (class starting points, cohort needs) — not in isolation.</p>
            <div className="space-y-2.5">
              {staff.map((s, i) => (
                <div key={s.name} className="flex items-center gap-2">
                  <div className="w-8 h-8 rounded-full text-white flex items-center justify-center text-xs font-black shrink-0" style={{ backgroundColor: BC.purple }}>{s.name.slice(0, 2).toUpperCase()}</div>
                  <div className="w-20">
                    <div className="text-sm font-bold" style={{ color: BC.ink }}>{s.name}</div>
                    <div className="text-[10px] text-slate-400">{s.pairs} unit{s.pairs === 1 ? "" : "s"}</div>
                  </div>
                  <div className="flex-1 h-3.5 rounded-full overflow-hidden" style={{ backgroundColor: "#ECE6F2" }}>
                    <div className="h-full rounded-full" style={{ width: Math.max((s.avg / maxU) * 100, 3) + "%", backgroundColor: i === 0 ? "#189E8A" : i === staff.length - 1 && staff.length > 1 ? "#FFAC38" : "#47ABFB" }}></div>
                  </div>
                  <span className="text-xs font-black w-12 text-right tabular-nums" style={{ color: BC.ink }}>+{s.avg.toFixed(2)}</span>
                  <span className="w-24 text-right">
                    {i === 0 ? <span className="text-[9px] font-black uppercase px-1.5 py-0.5 rounded" style={pill("#189E8A", "#fff")}>Top impact</span> : null}
                    {i === staff.length - 1 && staff.length > 1 ? <span className="text-[9px] font-black uppercase px-1.5 py-0.5 rounded" style={pill("#FFAC38", BC.ink)}>Dev focus</span> : null}
                  </span>
                </div>
              ))}
            </div>
          </div>

          <div className="bg-white rounded-2xl p-4 shadow-sm">
            <p className="fd font-bold text-sm mb-2" style={{ color: BC.ink }}>All units</p>
            <div className="space-y-1.5">
              {allPairs.map((p, i) => (
                <div key={i} className="flex items-center gap-2 text-xs">
                  <span className="font-bold px-1.5 py-0.5 rounded" style={pill(FRAMEWORK[p.pillar].hexL, BC.ink)}>{p.pillar.split(" ")[0]}</span>
                  <span className="flex-1 text-slate-600 truncate">{p.school.name} · {p.cls.name} · T{p.term} · Stg {p.stage}</span>
                  <span className="text-slate-400">{p.post ? p.post.educator : (p.pre ? p.pre.educator : "")}</span>
                  {p.uplift !== null ? <span className="font-black tabular-nums w-12 text-right" style={{ color: p.uplift >= 0.3 ? "#189E8A" : p.uplift >= 0.1 ? "#2554C7" : "#E06B22" }}>+{p.uplift.toFixed(2)}</span> : <span className="text-slate-300 font-bold w-12 text-right">—</span>}
                </div>
              ))}
            </div>
          </div>
          <p className="text-center text-[11px] text-slate-400 px-6">Production version: filter by academic year/term, export to PDF and Excel for annual company reporting, drill into any unit, and pillar-level breakdowns per staff member.</p>
        </div>
      </div>
    );
  }

  // ---------------- ADMIN CONSOLE ----------------
  if (screen === "admin") {
    const aSchool = schools.find(s => s.id === adminSchoolId);
    const aCls = aSchool ? aSchool.classes.find(c => c.id === adminClassId) : null;

    if (aSchool && aCls) return (
      <div className="fb min-h-screen" style={{ backgroundColor: BC.bg }}>
        {FONTS}
        <Header title={aCls.name} sub={aSchool.name + " · Admin"} back={() => setAdminClassId(null)} />
        <div className="max-w-lg mx-auto p-4 space-y-3">

          {/* Curriculum map editor */}
          <div className="bg-white rounded-2xl p-4 shadow-sm">
            <p className="fd font-bold text-sm mb-1 flex items-center gap-2" style={{ color: BC.ink }}><CalendarDays size={16} />Curriculum map</p>
            <p className="text-[11px] text-slate-400 mb-3">Physical Me runs every term (underpins PE). Set the paired pillar per term — usually agreed with the school during account review. Tap a paired pillar to change it; extend one across terms if progress shows children need longer.</p>
            <CurriculumMap map={aCls.map} current={CURRENT_TERM} onEdit={(term, which) => { if (which === "paired") setEditMap({ term }); }} />
            {editMap ? (
              <div className="mt-3 rounded-xl border p-3" style={{ borderColor: BC.mid }}>
                <p className="text-xs font-bold mb-2" style={{ color: BC.ink }}>{termLabel(editMap.term)} — paired pillar</p>
                <div className="grid grid-cols-3 gap-2">
                  {PAIRED_PILLARS.map(p => (
                    <button key={p} onClick={() => { setPaired(aSchool.id, aCls.id, editMap.term, p); setEditMap(null); }} className="py-2 rounded-lg text-xs font-bold text-white" style={{ backgroundColor: FRAMEWORK[p].hex, opacity: aCls.map.find(m => m.term === editMap.term).paired === p ? 1 : 0.6 }}>{p.split(" ")[0]}</button>
                  ))}
                </div>
                <button onClick={() => setEditMap(null)} className="w-full mt-2 py-1.5 rounded-lg text-xs font-bold" style={pill("#ECE6F2", "#64748b")}>Cancel</button>
              </div>
            ) : null}
          </div>

          <div className="bg-white rounded-2xl p-4 shadow-sm">
            <p className="fd font-bold text-sm mb-3 flex items-center gap-2" style={{ color: BC.ink }}><UserPlus size={16} />Add pupil</p>
            <div className="flex gap-2 mb-2">
              <input value={F("pName")} onChange={e => setF("pName", e.target.value)} placeholder="Full name" className="flex-1 px-3 py-2 rounded-lg border border-slate-300 text-sm" />
              <input value={F("pInit")} onChange={e => setF("pInit", e.target.value.toUpperCase())} placeholder="Initials" className="w-20 px-3 py-2 rounded-lg border border-slate-300 text-sm" />
            </div>
            <div className="flex gap-2 mb-2">
              <button onClick={() => setF("pG", F("pG", null) === "B" ? null : "B")} className="flex-1 py-1.5 rounded-lg text-xs font-bold border-2" style={F("pG", null) === "B" ? { backgroundColor: "#47ABFB", color: BC.ink, borderColor: "#47ABFB" } : { backgroundColor: BC.bg, color: "#94a3b8", borderColor: "#E5E0EB" }}>Boy</button>
              <button onClick={() => setF("pG", F("pG", null) === "G" ? null : "G")} className="flex-1 py-1.5 rounded-lg text-xs font-bold border-2" style={F("pG", null) === "G" ? { backgroundColor: "#FF6293", color: BC.ink, borderColor: "#FF6293" } : { backgroundColor: BC.bg, color: "#94a3b8", borderColor: "#E5E0EB" }}>Girl</button>
              {FLAGS.map(fl => {
                const fk = "p" + fl[1];
                const on = F(fk, false);
                return <button key={fl[0]} onClick={() => setF(fk, !on)} className="flex-1 py-1.5 rounded-lg text-xs font-bold border-2" style={on ? { backgroundColor: fl[2], color: fl[3], borderColor: fl[2] } : { backgroundColor: BC.bg, color: "#94a3b8", borderColor: "#E5E0EB" }}>{fl[1]}</button>;
              })}
            </div>
            <button onClick={() => addPupil(aSchool.id, aCls.id)} className="w-full py-2.5 rounded-xl font-bold text-sm hover:opacity-90" style={pill(BC.ink, "#fff")}>Add to register</button>
          </div>
          <div className="bg-white rounded-2xl p-4 shadow-sm">
            <p className="fd font-bold text-sm mb-3" style={{ color: BC.ink }}>{aCls.pupils.length} pupils on register</p>
            {aCls.pupils.length === 0 ? <p className="text-xs text-slate-400">No pupils yet. Add them above or use the Excel import on the school page. Registers are required for Level 3 per-pupil assessment and pupil-group impact breakdowns.</p> : null}
            <div className="space-y-1.5">
              {aCls.pupils.map(p => (
                <div key={p.init} className="flex items-center gap-2 text-sm">
                  <span className="w-9 text-[10px] font-black text-slate-400">{p.init}</span>
                  <span className="flex-1 font-medium truncate" style={{ color: BC.ink }}>{p.name}</span>
                  <FlagChips p={p} />
                  <button onClick={() => removePupil(aSchool.id, aCls.id, p.init)} className="p-1 text-slate-300 hover:text-rose-500"><Trash2 size={14} /></button>
                </div>
              ))}
            </div>
          </div>
        </div>
      </div>
    );

    if (aSchool) return (
      <div className="fb min-h-screen" style={{ backgroundColor: BC.bg }}>
        {FONTS}
        <Header title={aSchool.name} sub="Admin · Manage classes" back={() => { setAdminSchoolId(null); setImportPreview(null); }} />
        <div className="max-w-lg mx-auto p-4 space-y-3">
          <div className="bg-white rounded-2xl p-4 shadow-sm">
            <p className="fd font-bold text-sm mb-1 flex items-center gap-2" style={{ color: BC.ink }}><FileSpreadsheet size={16} />Import classes & pupils from Excel</p>
            <p className="text-[11px] text-slate-400 mb-3">Use the template so columns populate correctly. Accepts .xlsx or .csv. New classes get the default curriculum map, editable per class.</p>
            <div className="grid grid-cols-2 gap-2">
              <button onClick={downloadTemplate} className="py-2.5 rounded-xl border-2 font-bold text-xs flex items-center justify-center gap-1.5 hover:opacity-80" style={{ borderColor: BC.mid, color: BC.purple }}><Download size={14} />Download template</button>
              <label className="py-2.5 rounded-xl font-bold text-xs flex items-center justify-center gap-1.5 hover:opacity-90 cursor-pointer" style={pill(BC.ink, "#fff")}>
                <Upload size={14} />Upload file
                <input type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={e => handleImportFile(e, aSchool.id)} />
              </label>
            </div>
            {importPreview !== null && importPreview.sid === aSchool.id ? (
              <div className="mt-3 border-t border-slate-100 pt-3">
                {importPreview.list.length > 0 ? (
                  <div>
                    <p className="text-xs font-bold mb-2" style={{ color: BC.ink }}>Preview: {importPreview.list.length} class{importPreview.list.length === 1 ? "" : "es"}, {importPreview.list.reduce((n, c) => n + c.pupils.length, 0)} pupils</p>
                    {importPreview.list.map(c => (
                      <div key={c.name} className="text-xs text-slate-600 mb-1">
                        <b>{c.name}</b> ({c.year}, Stage {c.stage}) — {c.pupils.length} pupils · Boys {c.pupils.filter(p => p.g === "B").length} · Girls {c.pupils.filter(p => p.g === "G").length} · PP {c.pupils.filter(p => p.pp).length} · SEND {c.pupils.filter(p => p.send).length} · EAL {c.pupils.filter(p => p.eal).length}
                      </div>
                    ))}
                  </div>
                ) : null}
                {importPreview.errors.length > 0 ? (
                  <div className="mt-2 rounded-lg p-2 border" style={tintBox("#FFAC38")}>
                    {importPreview.errors.map((er, i) => <p key={i} className="text-[11px]" style={{ color: "#9a5b13" }}>{er}</p>)}
                  </div>
                ) : null}
                {importPreview.list.length > 0 ? (
                  <div className="grid grid-cols-2 gap-2 mt-3">
                    <button onClick={() => setImportPreview(null)} className="py-2 rounded-xl border-2 border-slate-200 text-slate-500 font-bold text-xs">Cancel</button>
                    <button onClick={confirmImport} className="py-2 rounded-xl font-bold text-xs hover:opacity-90" style={pill(BC.lime, BC.ink)}>Confirm import</button>
                  </div>
                ) : (
                  <button onClick={() => setImportPreview(null)} className="w-full mt-2 py-2 rounded-xl border-2 border-slate-200 text-slate-500 font-bold text-xs">Close</button>
                )}
              </div>
            ) : null}
          </div>

          <div className="bg-white rounded-2xl p-4 shadow-sm">
            <p className="fd font-bold text-sm mb-1 flex items-center gap-2" style={{ color: BC.ink }}><Image size={16} />School logo (for impact reports)</p>
            <div className="flex items-center gap-3 mt-2">
              {aSchool.logo ? <img src={aSchool.logo} alt="logo" className="h-12 w-12 object-contain rounded-lg border border-slate-200 bg-white" /> : <div className="h-12 w-12 rounded-lg text-white flex items-center justify-center font-black" style={{ backgroundColor: aSchool.colour }}>{aSchool.name[0]}</div>}
              <label className="flex-1 py-2.5 rounded-xl border-2 font-bold text-xs flex items-center justify-center gap-1.5 hover:opacity-80 cursor-pointer" style={{ borderColor: BC.mid, color: BC.purple }}>
                <Upload size={14} />{aSchool.logo ? "Replace logo" : "Upload logo"}
                <input type="file" accept="image/*" className="hidden" onChange={e => {
                  const f = e.target.files && e.target.files[0];
                  e.target.value = "";
                  if (!f) return;
                  const rd = new FileReader();
                  rd.onload = () => setLogo(aSchool.id, rd.result);
                  rd.readAsDataURL(f);
                }} />
              </label>
            </div>
          </div>

          <div className="bg-white rounded-2xl p-4 shadow-sm">
            <p className="fd font-bold text-sm mb-3 flex items-center gap-2" style={{ color: BC.ink }}><Plus size={16} />Add class manually</p>
            <div className="grid grid-cols-2 gap-2 mb-2">
              <input value={F("cName")} onChange={e => setF("cName", e.target.value)} placeholder="Class name" className="px-3 py-2 rounded-lg border border-slate-300 text-sm" />
              <select value={F("cYear", "Year 1")} onChange={e => setF("cYear", e.target.value)} className="px-3 py-2 rounded-lg border border-slate-300 text-sm bg-white">
                {["Reception", "Year 1", "Year 2", "Year 3", "Year 4", "Year 5", "Year 6"].map(y => <option key={y}>{y}</option>)}
              </select>
              <select value={F("cStage", 1)} onChange={e => setF("cStage", e.target.value)} className="px-3 py-2 rounded-lg border border-slate-300 text-sm bg-white">
                {[1, 2, 3, 4].map(s => <option key={s} value={s}>Default Stage {s}</option>)}
              </select>
              <input type="number" value={F("cSize", 30)} onChange={e => setF("cSize", e.target.value)} placeholder="Class size" className="px-3 py-2 rounded-lg border border-slate-300 text-sm" />
            </div>
            <p className="text-[11px] text-slate-400 mb-2">New classes start with the default map: Physical every term + Social / Mental / Personal rotating. Edit it after creating.</p>
            <button onClick={() => addClass(aSchool.id)} className="w-full py-2.5 rounded-xl font-bold text-sm hover:opacity-90" style={pill(BC.ink, "#fff")}>Add class</button>
          </div>
          {aSchool.classes.map(c => (
            <button key={c.id} onClick={() => { setAdminClassId(c.id); setEditMap(null); }} className="w-full bg-white rounded-2xl p-4 shadow-sm hover:shadow-md text-left">
              <div className="flex items-center gap-3 mb-2">
                <div className="w-11 h-11 rounded-xl flex items-center justify-center font-bold" style={pill(BC.lilac + "44", BC.purple)}>{c.name.slice(0, 2)}</div>
                <div className="flex-1">
                  <div className="font-bold" style={{ color: BC.ink }}>{c.name} <span className="font-normal text-slate-400">· {c.year}</span></div>
                  <div className="text-xs text-slate-500">{c.pupils.length} on register · size {c.size} · Stage {c.stage}</div>
                </div>
                <ChevronRight className="text-slate-300" />
              </div>
              <CurriculumMap map={c.map} current={CURRENT_TERM} />
            </button>
          ))}
        </div>
      </div>
    );

    return (
      <div className="fb min-h-screen" style={{ backgroundColor: BC.bg }}>
        {FONTS}
        <Header title="Admin Console" sub="Manage schools, classes and pupils" back={() => setScreen("schools")} />
        <div className="max-w-lg mx-auto p-4 space-y-3">
          <div className="bg-white rounded-2xl p-4 shadow-sm">
            <p className="fd font-bold text-sm mb-3 flex items-center gap-2" style={{ color: BC.ink }}><Building2 size={16} />Add school</p>
            <input value={F("sName")} onChange={e => setF("sName", e.target.value)} placeholder="School name" className="w-full px-3 py-2 rounded-lg border border-slate-300 text-sm mb-2" />
            <div className="grid grid-cols-3 gap-2 mb-3">
              {[1, 2, 3].map(l => (
                <button key={l} onClick={() => setF("sLevel", l)} className="py-2 rounded-lg text-xs font-bold border-2" style={Number(F("sLevel", 1)) === l ? { borderColor: BC.mid, backgroundColor: BC.lilac + "33", color: BC.purple } : { borderColor: "#E5E0EB", color: "#94a3b8" }}>Level {l}</button>
              ))}
            </div>
            <button onClick={addSchool} className="w-full py-2.5 rounded-xl font-bold text-sm hover:opacity-90" style={pill(BC.ink, "#fff")}>Create school</button>
            <p className="text-[11px] text-slate-400 mt-2">The level sets the data-capture tier for all classes in the school (pricing model).</p>
          </div>
          {schools.map(s => (
            <button key={s.id} onClick={() => setAdminSchoolId(s.id)} className="w-full bg-white rounded-2xl p-4 shadow-sm hover:shadow-md text-left flex items-center gap-4">
              <div className="w-12 h-12 rounded-xl text-white flex items-center justify-center font-black text-lg shrink-0" style={{ backgroundColor: s.colour }}>{s.name[0]}</div>
              <div className="flex-1">
                <div className="font-bold" style={{ color: BC.ink }}>{s.name}</div>
                <div className="text-xs text-slate-500">{s.classes.length} classes · {s.classes.reduce((n, c) => n + c.pupils.length, 0)} pupils on register</div>
                <div className="mt-1"><LevelBadge level={s.level} /></div>
              </div>
              <ChevronRight className="text-slate-300" />
            </button>
          ))}
          <p className="text-[11px] text-slate-400 text-center px-4">Educator accounts only see schools they're assigned to and cannot access the Admin Console or Company Dashboard.</p>
        </div>
      </div>
    );
  }

  // ---------------- SCHOOL VIEW ----------------
  if (screen === "school") {
    const pairs = buildPairs([school], assessments);
    const years = [...new Set(school.classes.map(c => c.year))];

    return (
      <div className="fb min-h-screen" style={{ backgroundColor: BC.bg }}>
        {FONTS}
        <Header title={school.name} sub={"Level " + school.level + " · " + termLabel(CURRENT_TERM)} back={() => setScreen("schools")} />
        <div className="max-w-lg mx-auto">
          <div className="flex p-2 gap-2">
            <button onClick={() => setTab("plan")} className="flex-1 flex items-center justify-center gap-1.5 py-2.5 rounded-xl font-bold text-xs" style={tab === "plan" ? pill(BC.ink, "#fff") : pill("#fff", "#94a3b8")}><CalendarDays size={15} />Plan</button>
            <button onClick={() => setTab("progress")} className="flex-1 flex items-center justify-center gap-1.5 py-2.5 rounded-xl font-bold text-xs" style={tab === "progress" ? pill(BC.ink, "#fff") : pill("#fff", "#94a3b8")}><Clock size={15} />Progress</button>
            <button onClick={() => setTab("reports")} className="flex-1 flex items-center justify-center gap-1.5 py-2.5 rounded-xl font-bold text-xs" style={tab === "reports" ? pill(BC.ink, "#fff") : pill("#fff", "#94a3b8")}><BarChart3 size={15} />Reports</button>
          </div>

          {/* ---------- PLAN: curriculum map drives logging ---------- */}
          {tab === "plan" ? (
            <div className="p-3 space-y-3">
              <div className="bg-white rounded-2xl p-3 shadow-sm">
                <p className="text-xs text-slate-500">Each class follows a curriculum map — <b style={{ color: BC.ink }}>Physical Me every term</b> plus a rotating pillar. The current term is highlighted; tap this term's pillars to log or check.</p>
              </div>
              {school.classes.map(c => {
                const blocked = school.level === 3 && c.pupils.length === 0;
                const nowPaired = c.map.find(m => m.term === CURRENT_TERM).paired;
                const termPillars = ["Physical Me", nowPaired];
                return (
                  <div key={c.id} className="bg-white rounded-2xl p-4 shadow-sm">
                    <div className="flex items-center gap-2 mb-2">
                      <div className="w-9 h-9 rounded-lg flex items-center justify-center text-xs font-bold" style={pill(BC.lilac + "44", BC.purple)}>{c.name.slice(0, 2)}</div>
                      <div className="flex-1">
                        <div className="fd font-bold text-sm" style={{ color: BC.ink }}>{c.name} <span className="fb font-normal text-slate-400 text-xs">· {c.year} · Stg {c.stage}</span></div>
                      </div>
                    </div>
                    <div className="mb-3"><CurriculumMap map={c.map} current={CURRENT_TERM} statusFn={(pillar, term) => unitStatus(assessments, school, c, pillar, term)} /></div>
                    {blocked ? <p className="text-[11px] font-semibold" style={{ color: "#E3225C" }}>Admin must add the pupil register before Level 3 logging.</p> : (
                      <div>
                        <p className="text-[11px] font-bold uppercase mb-1.5" style={{ color: BC.purple }}>{termLabel(CURRENT_TERM)} — log or check</p>
                        <div className="grid grid-cols-2 gap-2">
                          {termPillars.map(p => {
                            const st = unitStatus(assessments, school, c, p, CURRENT_TERM);
                            const label = st === "complete" ? "Complete" : st === "awaiting" ? "Do check" : "Baseline";
                            const win = st === "awaiting" ? "check" : "base";
                            return (
                              <button key={p} onClick={() => {
                                if (st === "complete") { setClassId(c.id); setReport({ classId: c.id, pillar: p, stage: c.stage, term: CURRENT_TERM }); setScreen("report"); }
                                else { setClassId(c.id); setCfg({ pillar: p, stage: c.stage, window: win, block: 6, term: CURRENT_TERM, date: today() }); setShowStatements(true); setScreen("setup"); }
                              }} className="rounded-xl p-2 text-left text-white" style={{ backgroundColor: FRAMEWORK[p].hex }}>
                                <div className="text-[11px] font-bold">{p}</div>
                                <div className="text-[10px] flex items-center gap-1 opacity-90">
                                  {st === "complete" ? <CheckCircle2 size={11} /> : st === "awaiting" ? <Clock size={11} /> : <Circle size={11} />}{label}
                                </div>
                              </button>
                            );
                          })}
                        </div>
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          ) : null}

          {/* ---------- PROGRESS: audit trail ---------- */}
          {tab === "progress" ? (
            <div className="p-3 space-y-3">
              <div className="bg-white rounded-2xl p-3 shadow-sm">
                <p className="text-xs text-slate-500">Every planned unit across the year — what's baselined, awaiting its check, or complete.</p>
                <div className="flex gap-3 mt-2 text-[11px]">
                  <span className="flex items-center gap-1" style={{ color: "#0e6e60" }}><CheckCircle2 size={13} />Complete</span>
                  <span className="flex items-center gap-1" style={{ color: "#E06B22" }}><Clock size={13} />Awaiting check</span>
                  <span className="flex items-center gap-1 text-slate-400"><Circle size={13} />Not started</span>
                </div>
              </div>
              {school.classes.map(c => (
                <div key={c.id} className="bg-white rounded-2xl p-4 shadow-sm">
                  <div className="flex items-center gap-2 mb-2">
                    <div className="w-8 h-8 rounded-lg flex items-center justify-center text-xs font-bold" style={pill(BC.lilac + "44", BC.purple)}>{c.name.slice(0, 2)}</div>
                    <div className="fd font-bold text-sm" style={{ color: BC.ink }}>{c.name} <span className="fb font-normal text-slate-400 text-xs">· {c.year}</span></div>
                  </div>
                  <div className="space-y-1.5">
                    {c.map.map(m => {
                      const rowPillars = ["Physical Me", m.paired];
                      return rowPillars.map(p => {
                        const st = unitStatus(assessments, school, c, p, m.term);
                        const icon = st === "complete" ? <CheckCircle2 size={15} style={{ color: "#0e6e60" }} /> : st === "awaiting" ? <Clock size={15} style={{ color: "#E06B22" }} /> : <Circle size={15} className="text-slate-300" />;
                        return (
                          <div key={m.term + p} className="flex items-center gap-2 rounded-lg p-1.5" style={{ backgroundColor: m.term === CURRENT_TERM ? BC.lime + "22" : BC.bg }}>
                            <span className="shrink-0">{icon}</span>
                            <span className="text-[10px] font-black w-6 text-slate-400">T{m.term}</span>
                            <PillarChip p={p} small={true} />
                            <span className="flex-1"></span>
                            {st === "complete" ? (
                              <button onClick={() => { setClassId(c.id); setReport({ classId: c.id, pillar: p, stage: c.stage, term: m.term }); setScreen("report"); }} className="text-[10px] font-bold px-2 py-1 rounded-lg" style={pill(BC.ink, "#fff")}>Report</button>
                            ) : st === "awaiting" ? (
                              <button onClick={() => { setClassId(c.id); setCfg({ pillar: p, stage: c.stage, window: "check", block: 6, term: m.term, date: today() }); setShowStatements(true); setScreen("setup"); }} className="text-[10px] font-bold px-2 py-1 rounded-lg" style={pill(BC.lime, BC.ink)}>Do check</button>
                            ) : (
                              <button onClick={() => { setClassId(c.id); setCfg({ pillar: p, stage: c.stage, window: "base", block: 6, term: m.term, date: today() }); setShowStatements(true); setScreen("setup"); }} className="text-[10px] font-bold px-2 py-1 rounded-lg border" style={{ borderColor: "#E5E0EB", color: "#94a3b8" }}>Baseline</button>
                            )}
                          </div>
                        );
                      });
                    })}
                  </div>
                </div>
              ))}
            </div>
          ) : null}

          {tab === "reports" ? (
            <div className="p-3 space-y-3">
              <div className="text-white rounded-2xl p-4 shadow-sm" style={{ backgroundColor: BC.ink }}>
                <p className="fd font-bold text-sm flex items-center gap-2"><FileText size={16} />Generate School Impact Report</p>
                <p className="text-[11px] mt-1 mb-3" style={{ color: BC.lilac }}>Branded, shareable report built from your template. Choose the scope:</p>
                <div className="grid grid-cols-3 gap-2">
                  <button onClick={() => { setDocScope({ type: "school" }); setScreen("doc"); }} className="py-2 rounded-lg text-xs font-bold hover:opacity-90" style={pill(BC.lime, BC.ink)}>Whole school</button>
                  <button onClick={() => setScopePick(scopePick === "year" ? null : "year")} className="py-2 rounded-lg text-xs font-bold" style={scopePick === "year" ? pill("#fff", BC.ink) : pill(BC.purple, "#fff")}>By year</button>
                  <button onClick={() => setScopePick(scopePick === "class" ? null : "class")} className="py-2 rounded-lg text-xs font-bold" style={scopePick === "class" ? pill("#fff", BC.ink) : pill(BC.purple, "#fff")}>By class</button>
                </div>
                {scopePick === "year" ? (
                  <div className="flex flex-wrap gap-2 mt-3">
                    {years.map(y => <button key={y} onClick={() => { setDocScope({ type: "year", year: y }); setScreen("doc"); }} className="px-3 py-1.5 rounded-full text-xs font-bold hover:opacity-90" style={pill(BC.lilac, BC.ink)}>{y}</button>)}
                  </div>
                ) : null}
                {scopePick === "class" ? (
                  <div className="flex flex-wrap gap-2 mt-3">
                    {school.classes.map(c => <button key={c.id} onClick={() => { setDocScope({ type: "class", classId: c.id }); setScreen("doc"); }} className="px-3 py-1.5 rounded-full text-xs font-bold hover:opacity-90" style={pill(BC.lilac, BC.ink)}>{c.name}</button>)}
                  </div>
                ) : null}
              </div>

              <p className="text-xs font-bold uppercase px-1" style={{ color: BC.purple }}>Completed units</p>
              {pairs.filter(p => p.pre && p.post).length === 0 ? <p className="text-sm text-slate-500 text-center py-6">No completed units yet. A report is available once a unit has both a Baseline and a Progress Check.</p> : null}
              {pairs.filter(p => p.pre && p.post).map((p, i) => (
                <button key={i} onClick={() => { setClassId(p.cls.id); setReport({ classId: p.cls.id, pillar: p.pillar, stage: p.stage, term: p.term }); setScreen("report"); }} className="w-full bg-white rounded-2xl p-4 shadow-sm hover:shadow-md text-left">
                  <div className="flex items-center justify-between">
                    <div>
                      <PillarChip p={p.pillar} />
                      <span className="text-xs font-bold text-slate-400 ml-2">T{p.term} · Stg {p.stage}</span>
                    </div>
                    <ChevronRight className="text-slate-300" />
                  </div>
                  <div className="font-bold mt-1.5" style={{ color: BC.ink }}>{p.cls.name} · {p.cls.year}</div>
                  <div className="text-xs text-slate-500 mt-0.5">Baseline {fmtDate(p.pre.date)} · Check {fmtDate(p.post.date)}</div>
                </button>
              ))}
            </div>
          ) : null}
        </div>
      </div>
    );
  }

  // ---------------- SETUP ----------------
  if (screen === "setup") {
    const stageObj = cfg.pillar ? FRAMEWORK[cfg.pillar].stages[cfg.stage] : null;
    return (
      <div className="fb min-h-screen pb-6" style={{ backgroundColor: BC.bg }}>
        {FONTS}
        <Header title={cls.name + " · " + cls.year} sub={school.name + " · " + termLabel(cfg.term)} back={() => setScreen("school")} />
        <div className="max-w-lg mx-auto p-4 space-y-4">
          <div className="rounded-xl p-3 flex items-center gap-2" style={tintBox(BC.mid)}>
            <CalendarDays size={16} style={{ color: BC.purple }} />
            <p className="text-xs" style={{ color: BC.ink }}>Logging for <b>{termLabel(cfg.term)}</b> — pillar set from the class curriculum map. Change below if you're catching up a different term.</p>
          </div>
          <div>
            <p className="text-xs font-bold uppercase mb-2" style={{ color: BC.purple }}>Pillar</p>
            <div className="grid grid-cols-2 gap-2">
              {PILLARS.map(p => {
                const sel = cfg.pillar === p;
                return (
                  <button key={p} onClick={() => setCfg({ ...cfg, pillar: p })} className="fd py-3 rounded-xl font-bold text-sm text-white" style={{ backgroundColor: FRAMEWORK[p].hex, opacity: sel ? 1 : 0.55, boxShadow: sel ? "0 0 0 3px " + BC.ink : "none" }}>{p}</button>
                );
              })}
            </div>
          </div>
          <div>
            <p className="text-xs font-bold uppercase mb-2" style={{ color: BC.purple }}>Focal stage</p>
            <div className="grid grid-cols-4 gap-2">
              {[1, 2, 3, 4].map(s => (
                <button key={s} onClick={() => setCfg({ ...cfg, stage: s })} className="fd py-2.5 rounded-xl font-bold" style={cfg.stage === s ? pill(BC.ink, "#fff") : pill("#fff", "#64748b")}>{s}</button>
              ))}
            </div>
            {stageObj ? (
              <div className="mt-2 rounded-xl border overflow-hidden" style={{ borderColor: FRAMEWORK[cfg.pillar].hex }}>
                <button onClick={() => setShowStatements(!showStatements)} className="w-full flex items-center justify-between px-3 py-2" style={{ backgroundColor: FRAMEWORK[cfg.pillar].hex + "1F" }}>
                  <span className="text-xs font-bold" style={{ color: BC.ink }}>Stage {cfg.stage} · {stageObj.focus} — the "I can…" statements you'll assess</span>
                  <ChevronDown size={16} style={{ color: BC.ink, transform: showStatements ? "rotate(180deg)" : "none", transition: "transform .15s" }} />
                </button>
                {showStatements ? (
                  <ul className="px-3 py-2 space-y-1 bg-white">
                    {stageObj.statements.map((st, j) => (
                      <li key={j} className="text-xs flex gap-2" style={{ color: BC.ink }}>
                        <span className="font-bold" style={{ color: FRAMEWORK[cfg.pillar].hex }}>{j + 1}.</span>{st}
                      </li>
                    ))}
                    <li className="text-[10px] text-slate-400 pt-1">Check these read right before you start — it avoids logging against the wrong stage.</li>
                  </ul>
                ) : null}
              </div>
            ) : null}
          </div>
          <div>
            <p className="text-xs font-bold uppercase mb-2" style={{ color: BC.purple }}>Check point</p>
            <div className="grid grid-cols-2 gap-2">
              <button onClick={() => setCfg({ ...cfg, window: "base" })} className="py-2.5 rounded-xl font-bold text-sm" style={cfg.window === "base" ? pill(BC.ink, "#fff") : pill("#fff", "#64748b")}>Baseline (start)</button>
              <button onClick={() => setCfg({ ...cfg, window: "check" })} className="py-2.5 rounded-xl font-bold text-sm" style={cfg.window === "check" ? pill(BC.ink, "#fff") : pill("#fff", "#64748b")}>Progress Check</button>
            </div>
          </div>
          {cfg.window === "base" ? (
            <div>
              <p className="text-xs font-bold uppercase mb-2" style={{ color: BC.purple }}>Block length</p>
              <div className="grid grid-cols-3 gap-2">
                {[[6, "6 weeks"], [12, "12 weeks"], [999, "Other"]].map(o => (
                  <button key={o[0]} onClick={() => setCfg({ ...cfg, block: o[0] })} className="py-2 rounded-xl font-bold text-sm" style={cfg.block === o[0] ? pill(BC.ink, "#fff") : pill("#fff", "#64748b")}>{o[1]}</button>
                ))}
              </div>
              <p className="text-[11px] text-slate-400 mt-1 px-1">Tells supervisors whether a progress check is genuinely due or just mid-block.</p>
            </div>
          ) : null}
          <div>
            <p className="text-xs font-bold uppercase mb-2" style={{ color: BC.purple }}>Date</p>
            <input type="date" value={cfg.date} onChange={e => setCfg({ ...cfg, date: e.target.value })} className="w-full p-2.5 rounded-xl border border-slate-300 bg-white" />
          </div>
          <button disabled={!cfg.pillar} onClick={() => { setDraft(newDraft(school.level)); setStTab(0); setScreen("entry"); }} className="fd w-full py-3.5 rounded-xl font-bold disabled:opacity-30 hover:opacity-90" style={pill(BC.lime, BC.ink)}>Start logging →</button>
        </div>
      </div>
    );
  }

  // ---------------- ENTRY ----------------
  if (screen === "entry") {
    const stage = FRAMEWORK[cfg.pillar].stages[cfg.stage];
    const lvl = school.level;
    const preStageN = Math.max(cfg.stage - 1, 1);
    const postStageN = Math.min(cfg.stage + 1, 4);
    return (
      <div className="fb min-h-screen pb-28" style={{ backgroundColor: BC.bg }}>
        {FONTS}
        <Header title={cfg.pillar + " · Stage " + cfg.stage} sub={cls.name + " · " + WLABEL[cfg.window] + " · T" + cfg.term + " · " + fmtDate(cfg.date)} back={() => setScreen("setup")} />
        <div className="max-w-lg mx-auto p-3 space-y-3">
          <div className="rounded-xl p-3 text-white" style={{ backgroundColor: FRAMEWORK[cfg.pillar].hex }}>
            <div className="text-xs font-semibold opacity-80 uppercase">Learning focus</div>
            <div className="fd font-bold text-lg">{stage.focus}</div>
          </div>

          {lvl === 1 ? stage.statements.map((st, j) => (
            <div key={j} className="bg-white rounded-2xl p-4 shadow-sm">
              <p className="font-semibold text-sm mb-3" style={{ color: BC.ink }}>{j + 1}. {st}</p>
              {RATINGS.map(r => {
                const others = RATINGS.filter(x => x !== r).reduce((s, x) => s + draft.statements[j][x], 0);
                const roomLeft = cls.size - others;
                return (
                  <div key={r} className="flex items-center justify-between py-1.5">
                    <RatingChip r={r} />
                    <Counter value={draft.statements[j][r]} max={roomLeft} onChange={v => { const d = JSON.parse(JSON.stringify(draft)); d.statements[j][r] = v; setDraft(d); }} />
                  </div>
                );
              })}
              <p className="text-xs mt-1 text-right" style={{ color: total(draft.statements[j]) >= cls.size ? "#0e6e60" : "#94a3b8" }}>{total(draft.statements[j])} / {cls.size} pupils evidenced{total(draft.statements[j]) >= cls.size ? " ✓ full class" : ""}</p>
            </div>
          )) : null}

          {lvl === 2 ? stage.statements.map((st, j) => (
            <div key={j} className="bg-white rounded-2xl p-4 shadow-sm">
              <p className="font-semibold text-sm mb-3" style={{ color: BC.ink }}>{j + 1}. {st}</p>
              <div className="grid grid-cols-3 gap-2">
                {RATINGS.map(r => {
                  const s = draft.statements[j];
                  const key = j + "-" + r;
                  const typed = initInput[key] || "";
                  return (
                    <div key={r} className="rounded-xl border-2 p-2" style={{ borderColor: s.many === r ? BC.ink : "#E5E0EB" }}>
                      <MiniRating r={r} />
                      <button onClick={() => { const d = JSON.parse(JSON.stringify(draft)); d.statements[j].many = d.statements[j].many === r ? null : r; setDraft(d); }} className="w-full mt-1.5 py-1 rounded-md text-xs font-extrabold" style={s.many === r ? pill(BC.ink, "#fff") : pill("#ECE6F2", "#94a3b8")}>MANY</button>
                      <input value={typed} onChange={e => setInitInput({ ...initInput, [key]: e.target.value.toUpperCase() })}
                        onKeyDown={e => { if (e.key === "Enter" && typed.trim()) { const d = JSON.parse(JSON.stringify(draft)); d.statements[j].initials[r].push(typed.trim()); setDraft(d); setInitInput({ ...initInput, [key]: "" }); } }}
                        placeholder="Initials ↵" className="w-full mt-1.5 px-1.5 py-1 text-xs rounded border border-slate-200 focus:outline-none" />
                      <InitialChips list={s.initials[r]} onRemove={k => { const d = JSON.parse(JSON.stringify(draft)); d.statements[j].initials[r].splice(k, 1); setDraft(d); }} />
                    </div>
                  );
                })}
              </div>
              <p className="text-[11px] text-slate-400 mt-2">Tap MANY where the majority sit; add initials for pupils above or below that level.</p>
            </div>
          )) : null}

          {lvl === 3 ? (
            <div className="bg-white rounded-2xl p-4 shadow-sm">
              <div className="flex gap-1.5 mb-2">
                {stage.statements.map((s, j) => (
                  <button key={j} onClick={() => setStTab(j)} className="flex-1 py-1.5 rounded-lg text-xs font-bold" style={stTab === j ? pill(BC.ink, "#fff") : pill("#ECE6F2", "#94a3b8")}>S{j + 1}</button>
                ))}
              </div>
              <p className="font-semibold text-sm mb-3" style={{ color: BC.ink }}>{stTab + 1}. {stage.statements[stTab]}</p>
              <div className="space-y-1.5">
                {cls.pupils.map(p => {
                  const cur = draft.pupils[p.init][stTab];
                  const flag = draft.flags[p.init];
                  return (
                    <div key={p.init} className="flex items-center gap-2">
                      <button onClick={() => { const d = JSON.parse(JSON.stringify(draft)); const order = [undefined, "below", "above"]; const nx = order[(order.indexOf(flag) + 1) % 3]; if (nx) { d.flags[p.init] = nx; } else { delete d.flags[p.init]; } setDraft(d); }}
                        title="Tap to flag working below/above stage"
                        className="w-7 h-7 rounded-md text-xs font-bold shrink-0"
                        style={flag === "below" ? pill("#FFAC38", BC.ink) : flag === "above" ? pill("#47FBDF", BC.ink) : pill("#ECE6F2", "#cbd5e1")}>{flag === "below" ? "↓" : flag === "above" ? "↑" : "–"}</button>
                      <div className="flex-1 min-w-0">
                        <span className="text-sm font-medium" style={{ color: BC.ink }}>{p.name} </span>
                        <FlagChips p={p} small={true} />
                      </div>
                      <div className="flex gap-1">
                        {RATINGS.map(r => (
                          <button key={r} onClick={() => { const d = JSON.parse(JSON.stringify(draft)); if (cur === r) { delete d.pupils[p.init][stTab]; } else { d.pupils[p.init][stTab] = r; } setDraft(d); }}
                            className="w-10 py-1.5 rounded-md text-xs font-bold border"
                            style={cur === r ? { backgroundColor: RCOL[r], color: RTXT[r], borderColor: RCOL[r] } : { backgroundColor: "#fff", color: "#94a3b8", borderColor: "#E5E0EB" }}>{r === "Es" ? "Est" : RLABEL[r][0]}</button>
                        ))}
                      </div>
                    </div>
                  );
                })}
              </div>
              <p className="text-[11px] text-slate-400 mt-2">↓ / ↑ flags a pupil working below or above the focal stage. Leave blank where no evidence was shown.</p>
            </div>
          ) : null}

          {lvl < 3 ? (
            <div className="bg-white rounded-2xl p-4 shadow-sm">
              <p className="font-bold text-sm mb-1" style={{ color: BC.ink }}>Working outside the focal stage</p>
              <p className="text-[11px] text-slate-400 mb-3">Capture pupils showing evidence below (Stage {preStageN}) or above (Stage {postStageN}) this stage.</p>
              {lvl === 1 ? (
                <div className="space-y-2">
                  <div className="flex items-center justify-between"><span className="text-sm text-slate-600">Pre-stage ({preStageN})</span><Counter value={draft.preCount} onChange={v => setDraft({ ...draft, preCount: v })} /></div>
                  <div className="flex items-center justify-between"><span className="text-sm text-slate-600">Post-stage ({postStageN})</span><Counter value={draft.postCount} onChange={v => setDraft({ ...draft, postCount: v })} /></div>
                </div>
              ) : (
                <div className="grid grid-cols-2 gap-2">
                  {[["pre", "Pre-stage (" + preStageN + ")"], ["post", "Post-stage (" + postStageN + ")"]].map(pair => {
                    const k = pair[0], t = pair[1];
                    const typed = initInput[k] || "";
                    return (
                      <div key={k} className="rounded-xl border-2 p-2" style={{ borderColor: "#E5E0EB" }}>
                        <div className="text-xs font-bold text-slate-500">{t}</div>
                        <input value={typed} onChange={e => setInitInput({ ...initInput, [k]: e.target.value.toUpperCase() })}
                          onKeyDown={e => { if (e.key === "Enter" && typed.trim()) { const d = JSON.parse(JSON.stringify(draft)); d[k].push(typed.trim()); setDraft(d); setInitInput({ ...initInput, [k]: "" }); } }}
                          placeholder="Initials ↵" className="w-full mt-1.5 px-1.5 py-1 text-xs rounded border border-slate-200" />
                        <InitialChips list={draft[k]} onRemove={i => { const d = JSON.parse(JSON.stringify(draft)); d[k].splice(i, 1); setDraft(d); }} />
                      </div>
                    );
                  })}
                </div>
              )}
            </div>
          ) : null}
        </div>
        <div className="fixed bottom-0 left-0 right-0 bg-white border-t border-slate-200 p-3">
          <div className="max-w-lg mx-auto">
            <button onClick={saveDraft} className="fd w-full py-3.5 rounded-xl font-bold hover:opacity-90" style={saved ? pill("#189E8A", "#fff") : pill(BC.lime, BC.ink)}>{saved ? "✓ Saved" : "Save " + WLABEL[cfg.window]}</button>
          </div>
        </div>
      </div>
    );
  }

  // ---------------- SCHOOL IMPACT REPORT DOCUMENT ----------------
  if (screen === "doc") {
    const allPairs = buildPairs([school], assessments);
    const scoped = allPairs.filter(p => {
      if (docScope.type === "school") return true;
      if (docScope.type === "year") return p.cls.year === docScope.year;
      return p.cls.id === docScope.classId;
    });
    const complete = scoped.filter(p => p.uplift !== null);
    const pending = scoped.filter(p => p.uplift === null);
    const scopeClasses = [...new Set(scoped.map(p => p.cls.id))].map(id => school.classes.find(c => c.id === id));
    const pupilsReached = scopeClasses.reduce((n, c) => n + c.size, 0);
    const avgUplift = complete.length ? complete.reduce((a, p) => a + p.uplift, 0) / complete.length : null;
    const aggAll = which => {
      const t = { E: 0, D: 0, Es: 0 };
      complete.forEach(p => { const a = aggRec(which === "pre" ? p.pre : p.post, p.cls); t.E += a.E; t.D += a.D; t.Es += a.Es; });
      return t;
    };
    const oPre = aggAll("pre"), oPost = aggAll("post");
    const estB = pct(oPre.Es, total(oPre)), estR = pct(oPost.Es, total(oPost));
    const emB = pct(oPre.E, total(oPre)), emR = pct(oPost.E, total(oPost));
    // Absolute average position (1=Emerging .. 3=Established) for the progress gauge —
    // a visual read of the headline figure that doesn't need a caption to interpret.
    const avgPreScore = meanScore(oPre) || 1;
    const avgPostScore = meanScore(oPost) || 1;
    const byPillar = PILLARS.map(pl => {
      const ps = complete.filter(p => p.pillar === pl);
      return { pillar: pl, n: ps.length, uplift: ps.length ? ps.reduce((a, p) => a + p.uplift, 0) / ps.length : null };
    }).filter(x => x.uplift !== null);
    const maxU = Math.max(0.01, ...byPillar.map(b => b.uplift));
    const stagePl = stagePlacement(scoped);
    const l3Complete = complete.filter(p => p.pre.level === 3);
    const groupRows = GROUPS.map(g => {
      const ga = groupAgg(l3Complete, g.filter);
      if (!ga.n) return null;
      return { label: g.label, bg: g.bg, fg: g.fg, n: ga.n, pB: pct(ga.pre.Es, total(ga.pre)), pR: pct(ga.post.Es, total(ga.post)) };
    }).filter(x => x !== null);
    // Narratives segmented by pillar/development area, rather than scattered per class unit.
    const narrativePillars = PILLARS.map(pl => {
      const items = scoped.filter(p => p.pillar === pl && p.post).map(p => ({
        p, c: p.cls, stg: FRAMEWORK[p.pillar].stages[p.stage],
        usePupils: school.level === 3 && p.cls.pupils.length > 0,
      }));
      return { pillar: pl, items };
    }).filter(x => x.items.length > 0);
    const scopeLabel = docScope.type === "school" ? "Whole-School Impact Report" : docScope.type === "year" ? docScope.year + " Impact Report" : (school.classes.find(c => c.id === docScope.classId).name + " Class Impact Report");
    const yearOrder = ["Reception", "Year 1", "Year 2", "Year 3", "Year 4", "Year 5", "Year 6"];
    const sortedClasses = [...scopeClasses].sort((a, b) => yearOrder.indexOf(a.year) - yearOrder.indexOf(b.year));

    return (
      <div className="fb min-h-screen py-6 px-3" style={{ backgroundColor: BC.purple }}>
        {FONTS}
        <div className="max-w-2xl mx-auto mb-3 flex gap-2">
          <button onClick={() => { setScreen("school"); setTab("reports"); }} className="flex-1 py-2.5 rounded-xl font-bold text-sm flex items-center justify-center gap-2 hover:opacity-90" style={pill(BC.ink, "#fff")}><ChevronLeft size={16} />Back</button>
          <button onClick={() => window.print()} className="flex-1 py-2.5 rounded-xl font-bold text-sm flex items-center justify-center gap-2 hover:opacity-90" style={pill(BC.lime, BC.ink)}><Printer size={16} />Print / save PDF</button>
        </div>
        <div className="max-w-2xl mx-auto bg-white shadow-xl rounded-sm p-8">
          <div className="flex items-center justify-between pb-4 mb-5" style={{ borderBottom: "4px solid " + BC.mid }}>
            <div className="flex items-center gap-3">
              {school.logo ? <img src={school.logo} alt="school logo" className="h-14 w-14 object-contain" /> : <div className="h-14 w-14 rounded-xl text-white flex items-center justify-center font-black text-xl" style={{ backgroundColor: school.colour }}>{school.name[0]}</div>}
              <div>
                <div className="fd font-bold text-lg leading-tight" style={{ color: BC.ink }}>{school.name}</div>
                <div className="text-xs text-slate-500">{docScope.type === "school" ? scopeClasses.length + " classes · " + pupilsReached + " pupils" : docScope.type === "year" ? docScope.year : ""}</div>
              </div>
            </div>
            <div className="flex flex-col items-end">
              <AMLogo s={44} light={false} />
              <div className="fd text-[9px] font-bold mt-1" style={{ color: BC.purple }}>activeme 360</div>
            </div>
          </div>

          <h1 className="fd text-2xl font-bold" style={{ color: BC.ink }}>{scopeLabel}</h1>
          <p className="text-sm text-slate-500 mt-1 mb-4">The ActiveMe Way · Children's Development Journey · Generated {fmtDate(today())}</p>

          <p className="text-sm leading-relaxed mb-5" style={{ color: "#1e1e1e" }}>
            At ActiveMe 360 our mission is to change lives by improving health and wellbeing through the power of physical activity, sport and education. Every child follows the <b>ActiveMe Way Children's Development Journey</b> — a structured pathway across <b>Physical, Mental, Social and Personal Me</b>, with age-appropriate "I can…" outcomes checked at the start (Baseline) and end (Progress Check) of each unit. This report evidences the measured impact of our partnership with {school.name}.
          </p>

          {complete.length > 0 ? (() => {
            const band = progressBand(avgUplift);
            return (
              <div className="rounded-2xl p-5 mb-4 text-white" style={{ backgroundColor: BC.ink }}>
                <div className="flex items-end justify-between gap-3 flex-wrap">
                  <div>
                    <p className="text-[10px] font-bold uppercase" style={{ color: BC.lilac }}>PE overall progress</p>
                    <div className="flex items-baseline gap-1.5 mt-1">
                      <div className="fd text-5xl font-bold leading-none">+{avgUplift.toFixed(2)}</div>
                      <div className="text-sm font-semibold" style={{ color: BC.lilac }}>/ +{MAX_UPLIFT.toFixed(2)} max</div>
                    </div>
                    <span className="inline-block mt-2 text-[10px] font-black uppercase px-2 py-1 rounded-full" style={{ backgroundColor: band.bg, color: band.fg }}>{band.label}</span>
                    <p className="text-[10px] mt-1 max-w-[220px]" style={{ color: "#E8DEF2" }}>{band.label} = {band.def}.</p>
                  </div>
                  <p className="text-sm font-semibold text-right leading-snug" style={{ color: "#F6F3F8" }}>The average pupil moved from<br /><b>{ratingWordAt(avgPreScore)}</b> to <b>{ratingWordAt(avgPostScore)}</b></p>
                </div>
                <div className="mt-4"><ProgressGauge preScore={avgPreScore} postScore={avgPostScore} light={true} /></div>
                <p className="text-[10px] mt-3" style={{ color: "#C9B8DA" }}>+{avgUplift.toFixed(2)} is {Math.round((avgUplift / MAX_UPLIFT) * 100)}% of the maximum possible improvement — every pupil moving a full level, Emerging → Established, on every statement. Combined across all {byPillar.length} pillar{byPillar.length === 1 ? "" : "s"} monitored this period — one number, not four separate scores — from {complete.length} completed unit{complete.length === 1 ? "" : "s"} across {pupilsReached} pupils.</p>
                <div className="mt-2 pt-2 border-t space-y-0.5" style={{ borderColor: "rgba(255,255,255,0.15)" }}>
                  <p className="text-[9px]" style={{ color: "#9C86AE" }}><b style={{ color: BC.lime }}>Excellent (1.00+):</b> children moving an average of 1 or more rating levels.</p>
                  <p className="text-[9px]" style={{ color: "#9C86AE" }}><b style={{ color: "#47FBDF" }}>Good (0.50–0.99):</b> children moving an average of 0.5 to just under 1 rating level.</p>
                  <p className="text-[9px]" style={{ color: "#9C86AE" }}><b style={{ color: "#FFD400" }}>Limited (under 0.50):</b> children moving an average of less than 0.5 of a rating level.</p>
                </div>
              </div>
            );
          })() : <p className="text-sm font-semibold mb-5" style={{ color: "#E06B22" }}>No completed units in this scope yet — sections below will populate as units complete their Progress Check.</p>}

          {complete.length > 0 ? (
            <div className="mb-1">
              <div className="grid grid-cols-3 gap-2">
                <div className="rounded-xl p-3 text-center border" style={tintBox(BC.mid)}>
                  <div className="text-[9px] font-bold uppercase" style={{ color: BC.purple }}>Units completed</div>
                  <div className="fd text-xl font-bold" style={{ color: BC.ink }}>{complete.length}</div>
                  <MetricNote>Baseline + Progress Check both logged</MetricNote>
                </div>
                <div className="rounded-xl p-3 text-center border" style={tintBox("#189E8A")}>
                  <div className="text-[9px] font-bold uppercase" style={{ color: "#0e6e60" }}>Established</div>
                  <div className="fd text-base font-bold" style={{ color: "#0e6e60" }}>{estB}%→{estR}%</div>
                  <MetricNote>{oPre.Es}→{oPost.Es} of {total(oPost)} ratings — shown confidently &amp; consistently</MetricNote>
                </div>
                <div className="rounded-xl p-3 text-center border" style={tintBox("#FFAC38")}>
                  <div className="text-[9px] font-bold uppercase" style={{ color: "#9a5b13" }}>Emerging</div>
                  <div className="fd text-base font-bold" style={{ color: "#9a5b13" }}>{emB}%→{emR}%</div>
                  <MetricNote>{oPre.E}→{oPost.E} of {total(oPost)} ratings — only just introduced</MetricNote>
                </div>
              </div>
              <p className="text-[10px] text-slate-400 mt-2 mb-5">Established / Emerging are shares of every statement rating recorded — each pupil contributes up to 4 ratings per unit, so this is not a headcount of pupils. Raw counts are shown above for small cohorts.</p>
            </div>
          ) : null}

          {byPillar.length > 0 ? (
            <div className="mb-6">
              <p className="fd text-xs font-bold uppercase mb-1" style={{ color: BC.purple }}>Impact by pillar</p>
              <p className="text-[11px] text-slate-400 mb-2">Average rating-level progress for each ActiveMe Way pillar monitored in this scope — read alongside the combined figure above to see which areas are driving it.</p>
              <div className="space-y-2">
                {byPillar.map(b => (
                  <div key={b.pillar} className="flex items-center gap-2">
                    <span className="w-24 text-[11px] font-bold px-2 py-1 rounded text-center" style={pill(FRAMEWORK[b.pillar].hexL, BC.ink)}>{b.pillar}</span>
                    <span className="text-[10px] text-slate-400 w-10">n={b.n}</span>
                    <div className="flex-1 h-3.5 rounded-full overflow-hidden" style={{ backgroundColor: "#ECE6F2" }}>
                      <div className="h-full rounded-full" style={{ width: Math.max((b.uplift / maxU) * 100, 3) + "%", backgroundColor: FRAMEWORK[b.pillar].hex }}></div>
                    </div>
                    <span className="text-xs font-black w-12 text-right tabular-nums" style={{ color: BC.ink }}>+{b.uplift.toFixed(2)}</span>
                  </div>
                ))}
              </div>
            </div>
          ) : null}

          {(stagePl.base.total > 0 || stagePl.check.total > 0) ? (
            <div className="mb-6">
              <p className="fd text-xs font-bold uppercase mb-1" style={{ color: BC.purple }}>Stage placement</p>
              <p className="text-[11px] text-slate-400 mb-2">% of pupils working below, at, or above their class's expected stage for this unit — Baseline vs Progress Check shows whether more children are on track by the end.</p>
              <StagePlacementBar label="Baseline" d={stagePl.base} />
              <StagePlacementBar label="Progress Check" d={stagePl.check} />
              <div className="flex gap-3 mt-1 text-[10px] text-slate-500 flex-wrap">
                <span className="flex items-center gap-1"><span className="w-2.5 h-2.5 rounded-sm" style={{ backgroundColor: "#FFAC38" }}></span>Below expected stage</span>
                <span className="flex items-center gap-1"><span className="w-2.5 h-2.5 rounded-sm" style={{ backgroundColor: "#189E8A" }}></span>At expected stage</span>
                <span className="flex items-center gap-1"><span className="w-2.5 h-2.5 rounded-sm" style={{ backgroundColor: "#47ABFB" }}></span>Above expected stage</span>
              </div>
            </div>
          ) : null}

          {school.level === 3 && groupRows.length > 0 ? (
            <div className="mb-6">
              <p className="fd text-xs font-bold uppercase mb-1" style={{ color: BC.purple }}>Impact by pupil group</p>
              <p className="text-[11px] text-slate-400 mb-2">% of statement ratings at Established, combined across every Level 3 pillar in this scope — dark = Baseline, light = Progress Check. “n=” is the number of pupils in that group. Demonstrates inclusive impact across boys, girls, Pupil Premium, SEND and EAL cohorts.</p>
              <GroupBars rows={groupRows} />
            </div>
          ) : null}
          {school.level < 3 ? (
            <div className="mb-6 rounded-xl p-4 border-2 border-dashed" style={{ backgroundColor: BC.bg, borderColor: BC.lilac }}>
              <p className="text-xs font-bold flex items-center gap-2" style={{ color: BC.purple }}><Lock size={13} />Pupil-group impact — Boys · Girls · PP · SEND · EAL</p>
              <p className="text-[11px] text-slate-500 mt-1">Cohort-level breakdowns require individual pupil assessment, available with a <b>Level 3 partnership</b>. Ask your ActiveMe 360 contact for a sample.</p>
            </div>
          ) : null}

          <div className="border-t border-slate-200 pt-5 mb-2">
            <p className="fd text-xs font-bold uppercase mb-1" style={{ color: BC.purple }}>{docScope.type === "class" ? "Unit detail" : "Class-by-class detail"}</p>
            <div className="flex gap-3 text-[10px] text-slate-500 flex-wrap">
              {RATINGS.map(r => <span key={r} className="flex items-center gap-1"><span className="w-2.5 h-2.5 rounded-sm" style={{ backgroundColor: RCOL[r] }}></span>{RLABEL[r]}</span>)}
              <span className="flex items-center gap-1"><span className="w-2.5 h-2.5 rounded-sm border border-slate-200" style={{ backgroundColor: "#ECE6F2" }}></span>Not yet evidenced</span>
            </div>
            <p className="text-[10px] text-slate-400 mt-1">Each bar below shows the number of pupils rated at each level for that statement (not a percentage) — read Base then Check, left to right.</p>
          </div>
          {sortedClasses.map(c => {
            const cPairs = scoped.filter(p => p.cls.id === c.id);
            return (
              <div key={c.id} className="mb-6">
                <div className="flex items-center gap-2 mb-2">
                  <div className="w-8 h-8 rounded-lg text-white flex items-center justify-center text-xs font-black" style={{ backgroundColor: BC.purple }}>{c.name.slice(0, 2)}</div>
                  <div>
                    <span className="fd font-bold" style={{ color: BC.ink }}>{c.name}</span>
                    <span className="text-xs text-slate-500"> · {c.year} · {c.size} pupils</span>
                  </div>
                </div>
                {cPairs.map((p, i) => {
                  const stg = FRAMEWORK[p.pillar].stages[p.stage];
                  const cb = p.pre ? aggRec(p.pre, c) : null;
                  const cr = p.post ? aggRec(p.post, c) : null;
                  const showPupilTable = school.level === 3 && c.pupils.length > 0 && p.pre && p.post;
                  return (
                    <div key={i} className="border rounded-xl p-3 mb-3" style={{ borderColor: "#E5E0EB", borderLeft: "4px solid " + FRAMEWORK[p.pillar].hex }}>
                      <div className="flex items-center justify-between flex-wrap gap-1">
                        <div>
                          <PillarChip p={p.pillar} small={true} />
                          <span className="text-[11px] font-bold text-slate-400 ml-2">T{p.term} · Stage {p.stage}: {stg.focus}</span>
                        </div>
                        {p.uplift !== null ? (
                          <span className="text-[11px] font-black" style={{ color: "#2554C7" }}>Avg progress +{p.uplift.toFixed(2)}</span>
                        ) : <span className="text-[11px] font-bold" style={{ color: "#E06B22" }}>Progress Check pending</span>}
                      </div>
                      <div className="text-[10px] text-slate-400 mt-0.5 mb-2">
                        {p.pre ? "Baseline " + fmtDate(p.pre.date) : ""}{p.post ? " · Check " + fmtDate(p.post.date) : ""}
                        {cb && cr ? " · Established " + pct(cb.Es, total(cb)) + "% → " + pct(cr.Es, total(cr)) + "% (" + cb.Es + "→" + cr.Es + " of " + total(cr) + " ratings)" : ""}
                      </div>
                      {stg.statements.map((st, j) => (
                        <div key={j} className="mb-2">
                          <p className="text-[11px] font-semibold mb-1" style={{ color: BC.ink }}>{j + 1}. {st}</p>
                          <div className="space-y-0.5">
                            {p.pre ? <div className="flex items-center gap-2"><span className="w-9 text-[9px] font-bold text-slate-400">Base</span><Bar d={dist(p.pre, j, c)} size={c.size} /></div> : null}
                            {p.post ? <div className="flex items-center gap-2"><span className="w-9 text-[9px] font-bold text-slate-400">Check</span><Bar d={dist(p.post, j, c)} size={c.size} /></div> : null}
                          </div>
                        </div>
                      ))}
                      {showPupilTable ? (
                        <div className="mt-3 pt-3 border-t" style={{ borderColor: "#F0EBF5" }}>
                          <p className="text-[11px] font-bold mb-1.5" style={{ color: BC.ink }}>Individual pupil journeys — {p.pillar} · Stage {p.stage}</p>
                          <PupilStatementTable pupils={c.pupils} stage={stg} b={p.pre} re={p.post} />
                        </div>
                      ) : null}
                    </div>
                  );
                })}
                {cPairs.length === 0 ? <p className="text-xs text-slate-400 mb-2">No units recorded for this class in the selected scope.</p> : null}
              </div>
            );
          })}
          {pending.length > 0 ? <p className="text-[11px] font-semibold mb-4" style={{ color: "#E06B22" }}>{pending.length} unit{pending.length === 1 ? " is" : "s are"} awaiting Progress Check and will be added automatically.</p> : null}

          <div className="mt-2 rounded-xl p-4 border" style={tintBox(BC.mid)}>
            <p className="fd text-xs font-bold uppercase mb-1" style={{ color: BC.purple }}>What this means for your school</p>
            <p className="text-sm leading-relaxed" style={{ color: "#1e1e1e" }}>
              These results show measurable development in the behaviours that underpin wellbeing, character and readiness to learn — not just activity delivered. They provide ready evidence for PE & Sport Premium reporting, Pupil Premium strategy reviews, governors and Ofsted's personal development judgement. The same "I can…" language is used with children in every session, so pupils can describe their own growth. We will agree next-term focus statements with class staff and recommend sharing highlights with parents and in celebration assemblies.
            </p>
          </div>

          {narrativePillars.length > 0 ? (
            <div className="mb-2 border-t border-slate-200 pt-5">
              <p className="fd text-xs font-bold uppercase mb-1 flex items-center gap-1.5" style={{ color: BC.purple }}>Impact narrative <span className="text-[9px] font-bold uppercase px-1.5 py-0.5 rounded" style={pill(BC.lime, BC.ink)}>Auto-drafted</span></p>
              <p className="text-[11px] text-slate-400 mb-3">Drafted automatically from the logged ratings and segmented by development area — review before sharing and edit if a statement needs more context.</p>
              {narrativePillars.map(np => (
                <div key={np.pillar} className="mb-4">
                  <div className="flex items-center gap-2 mb-2">
                    <span className="text-xs font-bold px-2 py-0.5 rounded-full" style={pill(FRAMEWORK[np.pillar].hexL, BC.ink)}>{np.pillar}</span>
                    <span className="fd font-bold text-sm" style={{ color: BC.ink }}>{PILLAR_OVERVIEW[np.pillar]}</span>
                  </div>
                  {np.items.map((it, i) => (
                    <div key={i} className="mb-3 pl-1">
                      <p className="text-[11px] font-bold text-slate-500 mb-1">{it.c.name} · {it.c.year} · Stage {it.p.stage}{it.p.term ? " · T" + it.p.term : ""}</p>
                      {it.usePupils ? (
                        <div className="space-y-1">
                          {it.c.pupils.map(pu => {
                            const t = pupilNarrative(pu, it.stg, it.p.pre, it.p.post);
                            return t ? <p key={pu.init} className="text-[11px] leading-relaxed" style={{ color: "#1e1e1e" }}><b>{pu.name}:</b> {t}</p> : null;
                          })}
                        </div>
                      ) : (
                        <p className="text-[11px] leading-relaxed" style={{ color: "#1e1e1e" }}>{classNarrative(it.stg, it.c, it.p.post)}</p>
                      )}
                    </div>
                  ))}
                </div>
              ))}
            </div>
          ) : null}

          <div className="mt-6 pt-4 border-t border-slate-200 flex items-center justify-between text-[11px] text-slate-400">
            <span>Prepared by ActiveMe 360 · {fmtDate(today())}</span>
            <span className="fd font-bold" style={{ color: BC.purple }}>www.activeme360.com</span>
          </div>
        </div>
        <p className="max-w-2xl mx-auto text-center text-[11px] mt-3" style={{ color: BC.lilac }}>Prototype with fictional data. Production version exports a polished branded PDF automatically, with term/year filters.</p>
      </div>
    );
  }

  // ---------------- INTERNAL REPORT ----------------
  if (screen === "report") {
    const rCls = school.classes.find(c => c.id === report.classId);
    const rTerm = report.term || 0;
    const recs = assessments.filter(a => a.schoolId === school.id && a.classId === report.classId && a.pillar === report.pillar && Number(a.stage) === Number(report.stage) && (a.term || 0) === rTerm);
    const b = recs.filter(r => r.window === "base").pop();
    const re = recs.filter(r => r.window === "check").pop();
    const stage = FRAMEWORK[report.pillar].stages[report.stage];
    const aB = b ? aggRec(b, rCls) : null;
    const aR = re ? aggRec(re, rCls) : null;
    const estB = aB ? pct(aB.Es, total(aB)) : 0;
    const estR = aR ? pct(aR.Es, total(aR)) : 0;
    const uplift = b && re ? pairUplift(b, re, rCls) : null;
    const preScore = aB ? (meanScore(aB) || 1) : 1;
    const postScore = aR ? (meanScore(aR) || 1) : 1;
    const groupRows = (school.level === 3 && b && re) ? GROUPS.map(g => {
      const ga = groupAgg([{ cls: rCls, pre: b, post: re }], g.filter);
      if (!ga.n) return null;
      return { label: g.label, bg: g.bg, fg: g.fg, n: ga.n, pB: pct(ga.pre.Es, total(ga.pre)), pR: pct(ga.post.Es, total(ga.post)) };
    }).filter(x => x !== null) : [];
    const stagePl = stagePlacement([{ cls: rCls, pre: b, post: re }]);

    const baselineOnly = b && !re;

    return (
      <div className="fb min-h-screen pb-10" style={{ backgroundColor: BC.bg }}>
        {FONTS}
        <Header title={baselineOnly ? "Baseline Pulse Check" : "Impact Report"} sub={rCls.name + " · " + rCls.year + " · " + school.name} back={() => { setScreen("school"); setTab(baselineOnly ? "progress" : "reports"); }} />
        <div className="max-w-lg mx-auto p-3 space-y-3">
          {baselineOnly ? (
            <div className="rounded-2xl p-3 border" style={tintBox("#FFAC38")}>
              <p className="text-xs font-bold flex items-center gap-1.5" style={{ color: "#9a5b13" }}><Clock size={14} />Baseline only — view for teaching, not yet a shareable report</p>
              <p className="text-[11px] mt-1" style={{ color: "#9a5b13" }}>This shows where the class started, useful for planning your teaching. The full impact report generates automatically once you complete the Progress Check.</p>
            </div>
          ) : null}

          <div className="bg-white rounded-2xl p-5 shadow-sm">
            <div className="flex items-center justify-between mb-1">
              <PillarChip p={report.pillar} />
              <LevelBadge level={school.level} />
            </div>
            <h2 className="fd text-xl font-bold" style={{ color: BC.ink }}>Stage {report.stage}: {stage.focus}</h2>
            <p className="text-xs text-slate-500 mt-1">
              {rTerm ? termLabel(rTerm) + " · " : ""}{b ? "Baseline " + fmtDate(b.date) + " (" + b.educator + ")" : "Baseline pending"}{re ? " → Progress Check " + fmtDate(re.date) + " (" + re.educator + ")" : baselineOnly ? " → Progress Check not yet done" : ""}
            </p>
            {b && re ? (
              <div className="mt-4">
                <div className="grid grid-cols-2 gap-2">
                  <div className="rounded-xl p-3 border" style={tintBox("#189E8A")}>
                    <div className="flex items-center gap-1" style={{ color: "#0e6e60" }}><TrendingUp size={14} /><span className="text-[10px] font-bold uppercase">Established</span></div>
                    <div className="fd text-2xl font-bold" style={{ color: "#0e6e60" }}>{estB}% → {estR}%</div>
                    <div className="text-[11px] font-semibold" style={{ color: "#0e6e60" }}>{estR - estB >= 0 ? "+" : ""}{estR - estB} pts across the unit</div>
                    <MetricNote>Share of ratings at Established ({aB.Es}→{aR.Es} of {total(aR)}) — children showing this confidently &amp; consistently.</MetricNote>
                  </div>
                  <div className="rounded-xl p-3 border" style={tintBox("#2554C7")}>
                    <div className="text-[10px] font-bold uppercase" style={{ color: "#2554C7" }}>Avg progress</div>
                    <div className="flex items-baseline gap-1">
                      <div className="fd text-2xl font-bold" style={{ color: "#2554C7" }}>{uplift !== null ? "+" + uplift.toFixed(2) : "—"}</div>
                      {uplift !== null ? <div className="text-[10px] font-semibold" style={{ color: "#2554C7" }}>/ +{MAX_UPLIFT.toFixed(2)} max</div> : null}
                    </div>
                    <div className="text-[11px] font-semibold" style={{ color: "#2554C7" }}>{ratingWordAt(preScore)} → {ratingWordAt(postScore)}</div>
                    {uplift !== null ? (() => {
                      const band = progressBand(uplift);
                      return (
                        <>
                          <span className="inline-block mt-1 text-[9px] font-black uppercase px-1.5 py-0.5 rounded-full" style={{ backgroundColor: band.bg, color: band.fg }}>{band.label}</span>
                          <MetricNote>{band.label} ({band.range}) = {band.def}.</MetricNote>
                        </>
                      );
                    })() : null}
                  </div>
                </div>
                <div className="rounded-xl p-3 border mt-2" style={tintBox(BC.mid)}>
                  <div className="text-[10px] font-bold uppercase mb-1" style={{ color: BC.purple }}>At a glance</div>
                  <ProgressGauge preScore={preScore} postScore={postScore} light={false} />
                </div>
              </div>
            ) : null}
          </div>

          {(stagePl.base.total > 0 || stagePl.check.total > 0) ? (
            <div className="bg-white rounded-2xl p-4 shadow-sm">
              <SectionTag pillar={report.pillar} stage={report.stage} term={rTerm} />
              <p className="fd font-bold text-sm mb-1" style={{ color: BC.ink }}>Stage placement</p>
              <p className="text-[11px] text-slate-400 mb-2">% of the class working below, at, or above Stage {report.stage} (their expected stage for this unit).</p>
              <StagePlacementBar label="Baseline" d={stagePl.base} />
              <StagePlacementBar label="Progress Check" d={stagePl.check} />
              <div className="flex gap-3 mt-1 text-[10px] text-slate-500 flex-wrap">
                <span className="flex items-center gap-1"><span className="w-2.5 h-2.5 rounded-sm" style={{ backgroundColor: "#FFAC38" }}></span>Below</span>
                <span className="flex items-center gap-1"><span className="w-2.5 h-2.5 rounded-sm" style={{ backgroundColor: "#189E8A" }}></span>At stage</span>
                <span className="flex items-center gap-1"><span className="w-2.5 h-2.5 rounded-sm" style={{ backgroundColor: "#47ABFB" }}></span>Above</span>
              </div>
            </div>
          ) : null}

          <div className="bg-white rounded-2xl p-4 shadow-sm">
            <SectionTag pillar={report.pillar} stage={report.stage} term={rTerm} />
            <p className="fd font-bold text-sm mb-3" style={{ color: BC.ink }}>{baselineOnly ? "Where the class is starting" : "Progress by \"I can…\" statement"}</p>
            <div className="flex gap-3 mb-1 text-[11px] text-slate-500 flex-wrap">
              {RATINGS.map(r => <span key={r} className="flex items-center gap-1"><span className="w-2.5 h-2.5 rounded-sm" style={{ backgroundColor: RCOL[r] }}></span>{RLABEL[r]}</span>)}
              <span className="flex items-center gap-1"><span className="w-2.5 h-2.5 rounded-sm border border-slate-200" style={{ backgroundColor: "#ECE6F2" }}></span>Not evidenced</span>
            </div>
            <p className="text-[10px] text-slate-400 mb-3">Each bar shows the number of pupils rated at each level for that statement (not a percentage).</p>
            {stage.statements.map((st, j) => (
              <div key={j} className="mb-4">
                <p className="text-xs font-semibold mb-1.5" style={{ color: BC.ink }}>{j + 1}. {st}</p>
                <div className="space-y-1">
                  {b ? <div className="flex items-center gap-2"><span className="w-9 text-[10px] font-bold text-slate-400">Base</span><Bar d={dist(b, j, rCls)} size={rCls.size} /></div> : null}
                  {re ? <div className="flex items-center gap-2"><span className="w-9 text-[10px] font-bold text-slate-400">Check</span><Bar d={dist(re, j, rCls)} size={rCls.size} /></div> : null}
                </div>
              </div>
            ))}
          </div>

          {groupRows.length > 0 ? (
            <div className="bg-white rounded-2xl p-4 shadow-sm">
              <SectionTag pillar={report.pillar} stage={report.stage} term={rTerm} />
              <p className="fd font-bold text-sm mb-1" style={{ color: BC.ink }}>Impact by pupil group</p>
              <p className="text-[11px] text-slate-400 mb-3">% of evidenced ratings at Established, Baseline → Progress Check. "n=" is the number of pupils in that group.</p>
              <GroupBars rows={groupRows} />
              <p className="text-[11px] text-slate-400 mt-3">Dark = Baseline · Light = Progress Check.</p>
            </div>
          ) : null}

          {school.level < 3 && !baselineOnly ? (
            <div className="rounded-2xl p-4 border-2 border-dashed" style={{ backgroundColor: "#ECE6F2", borderColor: BC.lilac }}>
              <div className="flex items-center gap-2 font-bold text-sm" style={{ color: BC.purple }}><Lock size={15} />Impact by pupil group — Boys · Girls · PP · SEND · EAL</div>
              <p className="text-xs text-slate-500 mt-1.5">Breakdowns require per-pupil assessment data. <b>Available on Level 3.</b></p>
            </div>
          ) : null}

          {school.level === 2 && re ? (
            <div className="bg-white rounded-2xl p-4 shadow-sm">
              <SectionTag pillar={report.pillar} stage={report.stage} term={rTerm} />
              <p className="fd font-bold text-sm mb-2" style={{ color: BC.ink }}>Named pupils to note (Progress Check)</p>
              {stage.statements.map((st, j) => {
                const s = re.data.statements[j];
                const any = RATINGS.some(r => s.initials[r].length > 0);
                if (!any) return null;
                return (
                  <div key={j} className="text-xs text-slate-600 mb-1.5">
                    <b>S{j + 1}:</b> {RATINGS.filter(r => s.initials[r].length > 0).map(r => RLABEL[r] + ": " + s.initials[r].join(", ")).join(" · ")}
                  </div>
                );
              })}
              {(re.data.pre.length > 0 || re.data.post.length > 0) ? (
                <div className="text-xs text-slate-600 mt-2 pt-2 border-t border-slate-100">
                  {re.data.pre.length > 0 ? <div><b>Working pre-stage:</b> {re.data.pre.join(", ")}</div> : null}
                  {re.data.post.length > 0 ? <div><b>Working post-stage:</b> {re.data.post.join(", ")}</div> : null}
                </div>
              ) : null}
            </div>
          ) : null}

          {school.level === 3 && b && re ? (
            <div className="bg-white rounded-2xl p-4 shadow-sm">
              <SectionTag pillar={report.pillar} stage={report.stage} term={rTerm} />
              <p className="fd font-bold text-sm mb-3" style={{ color: BC.ink }}>Individual pupil journeys</p>
              <PupilStatementTable pupils={rCls.pupils} stage={stage} b={b} re={re} />
            </div>
          ) : null}

          {!baselineOnly && re ? (
            <div className="bg-white rounded-2xl p-4 shadow-sm">
              <SectionTag pillar={report.pillar} stage={report.stage} term={rTerm} />
              <p className="fd font-bold text-sm mb-1 flex items-center gap-1.5" style={{ color: BC.ink }}>{PILLAR_OVERVIEW[report.pillar]} <span className="text-[9px] font-bold uppercase px-1.5 py-0.5 rounded" style={pill(BC.lime, BC.ink)}>Auto-drafted</span></p>
              <p className="text-[11px] text-slate-400 mb-2">Drafted automatically from the logged ratings — review before sharing and edit if a statement needs more context.</p>
              {school.level === 3 ? (
                <div className="space-y-1">
                  {rCls.pupils.map(p => {
                    const t = pupilNarrative(p, stage, b, re);
                    const flag = re.data.flags[p.init];
                    return t ? (
                      <p key={p.init} className="text-[11px] leading-relaxed" style={{ color: "#1e1e1e" }}>
                        <b>{p.name}</b>{flag === "above" ? <span className="ml-1 text-[10px] font-bold" style={{ color: "#0e6e60" }}>↑ above stage</span> : null}{flag === "below" ? <span className="ml-1 text-[10px] font-bold" style={{ color: "#9a5b13" }}>↓ below stage</span> : null}: {t}
                      </p>
                    ) : null;
                  })}
                </div>
              ) : (
                <p className="text-[11px] leading-relaxed" style={{ color: "#1e1e1e" }}>{classNarrative(stage, rCls, re)}</p>
              )}
            </div>
          ) : null}

          {!baselineOnly ? (
            <button onClick={() => { setDocScope({ type: "class", classId: report.classId }); setScreen("doc"); }} className="fd w-full py-3.5 rounded-xl font-bold flex items-center justify-center gap-2 hover:opacity-90" style={pill(BC.lime, BC.ink)}><FileText size={16} />Generate school impact report</button>
          ) : (
            <button onClick={() => { setClassId(report.classId); setCfg({ pillar: report.pillar, stage: report.stage, window: "check", block: (b && b.block) || 6, term: rTerm, date: today() }); setShowStatements(true); setScreen("setup"); }} className="fd w-full py-3.5 rounded-xl font-bold flex items-center justify-center gap-2 hover:opacity-90" style={pill(BC.lime, BC.ink)}><Clock size={16} />Complete the Progress Check</button>
          )}
          <p className="text-center text-[11px] text-slate-400 px-6">{baselineOnly ? "The shareable report unlocks once the Progress Check is logged." : "Or generate Whole school / Year group reports from the Reports tab."}</p>
        </div>
      </div>
    );
  }

  return null;
}

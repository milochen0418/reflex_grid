// The VBA editor's ears: the microphone to words. The AI Hub's Whisper hears (the socket its talking
// AI uses, /api/talk/stt: 16 kHz 16-bit PCM in, partial and final words out) once it is ready;
// until then, or without a hub, the browser's own speech recognition.

const WORKLET = `
class VbaPcm extends AudioWorkletProcessor {
  constructor() {
    super();
    this.step = sampleRate / 16000;
    this.pos = 0;
    this.sum = 0;
    this.count = 0;
    this.chunk = new Int16Array(1600);
    this.length = 0;
  }
  process(inputs) {
    const channel = inputs[0] && inputs[0][0];
    if (!channel) return true;
    for (let i = 0; i < channel.length; i++) {
      this.sum += channel[i];
      this.count += 1;
      this.pos += 1;
      if (this.pos < this.step) continue;
      this.pos -= this.step;
      const v = Math.max(-1, Math.min(1, this.sum / this.count));
      this.sum = 0;
      this.count = 0;
      this.chunk[this.length++] = v < 0 ? v * 0x8000 : v * 0x7fff;
      if (this.length === this.chunk.length) {
        this.port.postMessage(this.chunk.buffer, [this.chunk.buffer]);
        this.chunk = new Int16Array(1600);
        this.length = 0;
      }
    }
    return true;
  }
}
registerProcessor("vba-pcm", VbaPcm);
`;

const FIRST_STATUS_MS = 1200; // how long the hub gets to say it is there, before the browser listens

// What Whisper is told to expect: without it, 巨集 comes out as 聚集.
const HINT_ZH = "以下是用繁體中文討論 Excel VBA 程式：巨集、模組、工作表、儲存格、表單、按鈕、訊息方塊、迴圈、變數";
const HINT_EN = "Talking about Excel VBA code: macro, module, worksheet, range, cell, UserForm, button, MsgBox, loop, variable";

// The words that answer the change on screen, when they are all that was said.
const COMMANDS = {
  accept: ["接受", "好", "好的", "可以", "確定", "确定", "套用", "採用", "采用", "就這樣", "就这样", "accept", "ok", "okay", "yes", "apply"],
  reject: ["拒絕", "拒绝", "不要", "取消", "算了", "停", "停止", "reject", "cancel", "no", "stop", "discard"],
  fix: ["修正", "修一下", "修復", "修复", "修正錯誤", "修正错误", "幫我修", "帮我修", "fix", "fixit", "fixthis"],
};

export function commandOf(words) {
  const said = String(words || "").toLowerCase().replace(/[\s\p{P}]+/gu, "");
  for (const [command, list] of Object.entries(COMMANDS)) if (list.includes(said)) return command;
  return "";
}

export class Ears {
  // hub: the AI Hub's address ("" for none); hint(): the workbook's own words; onHearing(words so
  // far), onWords(a finished sentence), onNote(a word for the person, "" when all is well).
  constructor({ hub, hint, onHearing, onWords, onNote }) {
    this.hub = hub;
    this.hint = hint;
    this.onHearing = onHearing;
    this.onWords = onWords;
    this.onNote = onNote;
    this.lang = navigator.language || "en-US";
    this.on = false;
    this.mode = "";
    this.hubReady = false;
    this.timers = [];
  }

  async start() {
    if (this.on) return true;
    try {
      this.stream = await navigator.mediaDevices.getUserMedia({ audio: { echoCancellation: true, noiseSuppression: true, autoGainControl: true } });
    } catch (err) {
      this.onNote("The microphone can't be used here: " + (err && err.message ? err.message : err));
      return false;
    }
    this.on = true;
    if (this.hub) {
      this.connect();
      this.timers.push(setTimeout(() => this.choose(), FIRST_STATUS_MS));
    } else {
      this.choose();
    }
    return true;
  }

  stop() {
    this.on = false;
    this.timers.forEach(clearTimeout);
    this.timers = [];
    if (this.ws) {
      const ws = this.ws;
      this.ws = null;
      try { ws.close(); } catch (err) { /* closed */ }
    }
    this.stopBrowser();
    this.stopPcm();
    if (this.stream) {
      this.stream.getTracks().forEach((t) => t.stop());
      this.stream = null;
    }
    this.mode = "";
    this.hubReady = false;
    this.onHearing("");
  }

  connect() {
    if (!this.on) return;
    const lang = /^zh/i.test(this.lang) ? this.lang : ""; // else Whisper tells the language by itself
    const known = this.hint ? this.hint() : "";
    const hint = (lang ? HINT_ZH : HINT_EN) + (known ? ", " + known : "");
    const ws = new WebSocket(this.hub.replace(/^http/, "ws").replace(/\/$/, "") + "/api/talk/stt?lang=" + encodeURIComponent(lang) + "&hint=" + encodeURIComponent(hint));
    this.ws = ws;
    let opened = false;
    ws.onopen = () => {
      opened = true;
    };
    ws.onmessage = (ev) => {
      let m;
      try { m = JSON.parse(ev.data); } catch (err) { return; }
      if (m.type === "status") {
        this.hubReady = m.phase === "ready";
        this.phase = m.phase;
        this.choose();
      } else if (m.type === "subtitle" && this.mode === "hub") {
        if (!m.final) this.onHearing(m.text);
        else {
          this.onHearing("");
          if (m.text) this.onWords(m.text);
        }
      }
    };
    ws.onclose = () => {
      if (this.ws !== ws) return;
      this.hubReady = false;
      this.choose();
      if (this.on) this.timers.push(setTimeout(() => this.connect(), opened ? 2000 : 15000));
    };
  }

  // The hub's ears when they are ready, else the browser's; switched when that changes.
  choose() {
    if (!this.on) return;
    const mode = this.hubReady ? "hub" : "browser";
    if (mode === this.mode) return;
    if (mode === "hub") {
      this.stopBrowser();
      this.mode = "hub";
      this.onNote("");
      this.startPcm();
    } else {
      this.stopPcm();
      this.mode = "browser";
      this.startBrowser();
    }
  }

  async startPcm() {
    try {
      const Ctx = window.AudioContext || window.webkitAudioContext;
      const ctx = new Ctx();
      this.ctx = ctx;
      const url = URL.createObjectURL(new Blob([WORKLET], { type: "application/javascript" }));
      await ctx.audioWorklet.addModule(url);
      URL.revokeObjectURL(url);
      if (this.ctx !== ctx || !this.stream) return;
      const node = new AudioWorkletNode(ctx, "vba-pcm");
      const mute = ctx.createGain();
      mute.gain.value = 0;
      ctx.createMediaStreamSource(this.stream).connect(node);
      node.connect(mute).connect(ctx.destination);
      node.port.onmessage = (ev) => {
        if (this.mode === "hub" && this.ws && this.ws.readyState === 1) this.ws.send(ev.data);
      };
      this.node = node;
      if (ctx.state === "suspended") await ctx.resume();
    } catch (err) {
      this.onNote("Can't read the microphone: " + err.message);
    }
  }

  stopPcm() {
    if (this.node) {
      this.node.port.onmessage = null;
      this.node.disconnect();
      this.node = null;
    }
    if (this.ctx) {
      this.ctx.close().catch(() => {});
      this.ctx = null;
    }
  }

  startBrowser() {
    const R = window.SpeechRecognition || window.webkitSpeechRecognition;
    const waiting = this.phase === "downloading" || this.phase === "loading" || this.phase === "off";
    if (!R) {
      this.onNote(waiting ? "The AI Hub's ears are getting ready…" : "This browser can't hear, and the AI Hub's ears aren't there: type instead.");
      return;
    }
    this.onNote(waiting ? "The AI Hub's ears are getting ready: the browser listens meanwhile." : "");
    const rec = new R();
    rec.lang = /^zh/i.test(this.lang) ? (/^zh-(cn|hans)/i.test(this.lang) ? "zh-CN" : "zh-TW") : this.lang;
    rec.continuous = true;
    rec.interimResults = true;
    rec.onresult = (ev) => {
      let final = "";
      let interim = "";
      for (let i = ev.resultIndex; i < ev.results.length; i++) {
        const t = ev.results[i][0].transcript;
        if (ev.results[i].isFinal) final += t;
        else interim += t;
      }
      if (interim) this.onHearing(interim);
      if (final.trim()) {
        this.onHearing("");
        this.onWords(final.trim());
      }
    };
    rec.onend = () => {
      if (this.on && this.rec === rec) {
        try { rec.start(); } catch (err) { /* already */ }
      }
    };
    rec.onerror = (ev) => {
      if (ev.error === "not-allowed") this.onNote("The microphone isn't allowed here.");
    };
    this.rec = rec;
    try { rec.start(); } catch (err) { this.onNote("Can't listen: " + err.message); }
  }

  stopBrowser() {
    if (!this.rec) return;
    const rec = this.rec;
    this.rec = null;
    rec.onend = null;
    try { rec.stop(); } catch (err) { /* stopped */ }
  }
}

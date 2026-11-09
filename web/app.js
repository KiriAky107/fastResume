"use strict";

const byId = (id) => document.getElementById(id);
const editor = byId("yaml-editor");
const form = byId("basics-form");
const state = { resume: null, canonical: "", valid: false, busy: false, serial: 0, timer: null, presets: [] };
const PHOTO_PATTERN = /^data:image\/(?:png|jpeg|webp);base64,[A-Za-z0-9+/]+={0,2}$/;

function setPhoto(image, uri, focus) {
  const point = Array.isArray(focus) && focus.length === 2 && focus.every(Number.isFinite) ? focus : [0.5, 0.5];
  const position = () => {
    const width = image.naturalWidth;
    const height = image.naturalHeight;
    if (!width || !height) return;
    const visibleWidth = Math.min(width, height * 3 / 4);
    const visibleHeight = Math.min(height, width / (3 / 4));
    const align = (fraction, size, visible) => size === visible ? 0.5 : Math.max(0, Math.min(1, (fraction * size - visible / 2) / (size - visible)));
    image.style.objectPosition = (align(point[0], width, visibleWidth) * 100) + "% " + (align(point[1], height, visibleHeight) * 100) + "%";
  };
  image.onload = position;
  image.src = uri;
  if (image.complete) position();
}
const icons = {
  profile: ["M12 12a4 4 0 1 0 0-8 4 4 0 0 0 0 8", "M4 21v-2a8 8 0 0 1 16 0v2"],
  education: ["M2 8l10-5 10 5-10 5z", "M6 10v7c4 3 8 3 12 0v-7", "M22 8v8"],
  skills: ["M8 5L2 12l6 7", "M16 5l6 7-6 7", "M14 3l-4 18"],
  work: ["M3 7h18v14H3z", "M8 7V3h8v4", "M3 12h18", "M10 12v3h4v-3"],
  awards: ["M12 2l3 6 7 1-5 5 1 8-6-4-6 4 1-8-5-5 7-1z"],
  mail: ["M3 5h18v14H3z", "M3 5l9 8 9-8"],
  phone: ["M7 2h10v20H7z", "M10 18h4"],
  location: ["M20 10c0 6-8 12-8 12S4 16 4 10a8 8 0 0 1 16 0", "M12 7a3 3 0 1 0 0 6 3 3 0 0 0 0-6"],
  status: ["M4 5h16v16H4z", "M4 9h16", "M8 2v6", "M16 2v6", "M8 15l3 3 5-6"],
  globe: ["M12 2a10 10 0 1 0 0 20 10 10 0 0 0 0-20", "M2 12h20", "M12 2c6 5 6 15 0 20-6-5-6-15 0-20"],
};

function element(tag, className, text) {
  const node = document.createElement(tag);
  if (className) node.className = className;
  if (text !== undefined) node.textContent = text;
  return node;
}

function icon(name) {
  const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  svg.setAttribute("viewBox", "0 0 24 24");
  svg.setAttribute("fill", "none");
  svg.setAttribute("stroke", "currentColor");
  svg.setAttribute("stroke-width", "1.7");
  svg.setAttribute("stroke-linecap", "round");
  svg.setAttribute("stroke-linejoin", "round");
  svg.setAttribute("aria-hidden", "true");
  for (const path of icons[name] || icons.globe) {
    const item = document.createElementNS(svg.namespaceURI, "path");
    item.setAttribute("d", path);
    svg.append(item);
  }
  return svg;
}

function rich(parent, text, prefix = false) {
  text = String(text || "");
  if (prefix && !text.includes("**")) {
    const match = text.match(/^([^：\n]{1,20}：)([\s\S]*)$/);
    if (match) {
      parent.append(element("strong", "", match[1]), document.createTextNode(match[2]));
      return;
    }
  }
  const parts = text.split(/\*\*([\s\S]*?)\*\*/);
  parts.forEach((part, index) => {
    if (part) parent.append(index % 2 ? element("strong", "", part) : document.createTextNode(part));
  });
}

function textParagraph(text, className = "") {
  const paragraph = element("p", className);
  rich(paragraph, text);
  return paragraph;
}

function safeLink(label, url, className = "") {
  if (!url || !/^(https?:|mailto:|tel:)/i.test(url)) return element("span", className, label);
  const link = element("a", className, label);
  link.href = url;
  link.target = "_blank";
  link.rel = "noopener noreferrer";
  return link;
}

function section(paper, title, name) {
  const container = element("section", "resume-section");
  const heading = element("h3", "section-heading");
  const tile = element("span", "section-icon");
  tile.append(icon(name));
  heading.append(tile, document.createTextNode(title));
  const body = element("div", "section-body");
  container.append(heading, body);
  paper.append(container);
  return body;
}

function tint(color, white) {
  return "#" + [1, 3, 5].map((index) => {
    const value = parseInt(color.slice(index, index + 2), 16);
    return Math.round(value + (255 - value) * white).toString(16).padStart(2, "0");
  }).join("");
}

function render(data) {
  const paper = byId("preview");
  paper.replaceChildren();
  const theme = data.theme;
  paper.style.setProperty("--body-size", (theme.body_size * 4 / 3) + "px");
  paper.style.setProperty("--body-line", (theme.line_height * 4 / 3) + "px");
  paper.style.fontFamily = JSON.stringify(theme.font) + "," + JSON.stringify(theme.east_asia_font) + ",sans-serif";
  for (const [key, value] of Object.entries({ accent: theme.accent, ink: theme.ink, muted: theme.muted, tint: theme.background, border: theme.border, card: tint(theme.background, 0.88), skill: tint(theme.background, 0.5) })) {
    paper.style.setProperty("--" + key, value);
  }
  const basics = data.basics;
  const hero = element("header", "resume-hero");
  if (basics.photo && PHOTO_PATTERN.test(basics.photo)) {
    hero.classList.add("has-photo");
    const portrait = element("div", "hero-photo");
    const picture = element("img");
    setPhoto(picture, basics.photo, basics.photo_focus);
    picture.alt = basics.name + "的证件照";
    portrait.append(picture);
    hero.append(portrait);
  }
  const main = element("div", "hero-main");
  if (basics.eyebrow) main.append(element("div", "hero-eyebrow", basics.eyebrow));
  main.append(element("h2", "hero-name", basics.name));
  if (basics.position) main.append(element("p", "hero-position", basics.position));
  if (basics.subtitle) main.append(element("div", "hero-subtitle", basics.subtitle));
  const contacts = element("div", "hero-contact");
  for (const key of ["phone", "email"]) {
    if (!basics[key]) continue;
    const row = element("div", "contact-line");
    row.append(icon(key === "email" ? "mail" : key));
    const url = key === "email" ? "mailto:" + basics[key] : "tel:" + basics[key];
    row.append(safeLink(basics[key], url));
    contacts.append(row);
  }
  for (const item of data.code) {
    const row = element("div", "contact-line");
    row.append(icon(item.icon || "globe"), safeLink(item.label || item.url, item.url));
    contacts.append(row);
  }
  for (const key of ["location", "status"]) {
    if (!basics[key]) continue;
    const row = element("div", "contact-line contact-emphasis");
    row.append(icon(key), element("span", "", basics[key]));
    contacts.append(row);
  }
  if (basics.details) contacts.append(element("div", "contact-details", basics.details));
  if (basics.tagline) contacts.append(textParagraph(basics.tagline, "resume-tagline"));
  if (!contacts.childElementCount) contacts.append(element("div", "", ""));
  hero.append(main, contacts);
  paper.append(hero);
  const titles = data.section_titles;
  if (data.profile) section(paper, titles.profile, "profile").append(textParagraph(data.profile));
  const education = data.education.filter((item) => Object.values(item).some(Boolean));
  if (education.length) {
    const body = section(paper, titles.education, "education");
    for (const item of education) {
      const hasRow = Boolean(item.date || item.school || item.major);
      if (hasRow) {
        const row = element("div", "education-row");
        row.append(element("span", "education-date", item.date), element("span", "education-school", item.school), element("span", "education-major", item.major));
        body.append(row);
      }
      if (item.detail) body.append(textParagraph(item.detail, "education-detail" + (hasRow ? "" : " unindented")));
    }
  }
  const skills = data.skills;
  const groups = skills.groups.filter((item) => item.label || item.value);
  if (groups.length || skills.summary || skills.items.some(Boolean)) {
    const body = section(paper, titles.skills, "skills");
    if (skills.summary) body.append(textParagraph(skills.summary));
    if (groups.length) {
      const grid = element("div", "skill-grid");
      for (const item of groups) {
        const card = element("div", "skill-card");
        if (item.label) card.append(element("div", "skill-label", item.label));
        if (item.value) {
          const value = element("div", "skill-value");
          rich(value, item.value);
          card.append(value);
        }
        grid.append(card);
      }
      body.append(grid);
    }
    if (skills.items.length) {
      const list = element("ul", "plain-list");
      for (const item of skills.items) { const row = element("li"); rich(row, item); list.append(row); }
      body.append(list);
    }
  }
  for (const key of ["internship_experiences", "work_experiences", "project_experiences"]) {
    const records = data[key] || [];
    if (!records.length) continue;
    const body = section(paper, titles[key], "work");
    for (const item of records) {
      const card = element("div", "experience-card");
      if (item.name || item.date || item.url) {
        const heading = element("div", "experience-title");
        if (item.name || item.url) heading.append(safeLink((item.name || item.url) + (item.url ? " ↗" : ""), item.url, "experience-name"));
        if (item.date) heading.append(element("span", "experience-date", item.date));
        card.append(heading);
      }
      if (item.role) card.append(element("div", "experience-role", item.role));
      if (item.description) {
        const description = element("p", "experience-description");
        const label = key === "project_experiences" ? "项目简述" : key === "internship_experiences" ? "实习内容" : "工作内容";
        description.append(element("strong", "inline-label", label + "  "));
        rich(description, item.description);
        card.append(description);
      }
      if (item.highlights.length) {
        const list = element("ul");
        for (const text of item.highlights) { const row = element("li"); rich(row, text, true); list.append(row); }
        card.append(list);
      }
      if (item.tech_stack) {
        const stack = element("div", "experience-tech");
        stack.append(element("strong", "inline-label", "技术栈  "), document.createTextNode(item.tech_stack));
        card.append(stack);
      }
      body.append(card);
    }
  }
  const awards = data.certificates.filter((item) => typeof item === "string" ? Boolean(item) : Object.values(item).some(Boolean));
  if (awards.length) {
    const table = element("table", "award-table");
    const tbody = element("tbody");
    for (const item of awards) {
      const row = element("tr");
      if (typeof item === "string" || !item.name || !item.result) {
        const resultOnly = typeof item !== "string" && !item.name && Boolean(item.result);
        const cell = element("td", "award-single" + (resultOnly ? " award-result" : ""), typeof item === "string" ? item : item.name || item.result);
        cell.colSpan = 2;
        row.append(cell);
      } else row.append(element("td", "", item.name), element("td", "", item.result));
      tbody.append(row);
    }
    table.append(tbody);
    section(paper, titles.certificates, "awards").append(table);
  }
  for (const item of data.extra_sections) {
    if (!item.items.some(Boolean)) continue;
    const body = section(paper, item.title || "补充信息", "globe");
    for (const text of item.items) body.append(textParagraph(text));
  }
  byId("preview-name").textContent = theme.paper + " · " + basics.name;
}

function updateButtons() {
  for (const id of ["save-config", "export-docx", "export-pdf"]) byId(id).disabled = !state.valid || state.busy;
  byId("preset").disabled = !state.presets.length || state.busy;
  byId("basic-fields").disabled = !state.resume || state.busy;
  byId("import-button").disabled = state.busy;
  byId("format-button").disabled = state.busy;
  byId("upload-photo").disabled = !state.valid || state.busy;
  byId("remove-photo").disabled = !state.valid || state.busy || !state.resume?.basics.photo;
  editor.disabled = state.busy;
  for (const button of document.querySelectorAll("[data-add]")) button.disabled = !state.valid || state.busy;
}

function status(text, pending = false) {
  byId("status").textContent = text;
  byId("status").classList.toggle("pending", pending);
}

function error(text = "") {
  byId("error").textContent = text;
  byId("error").hidden = !text;
}

function fillForm(data) {
  for (const field of form.elements) {
    if (field.name) field.value = field.name === "profile" ? data.profile : (data.basics[field.name] || "");
  }
  const thumbnail = byId("photo-thumbnail");
  const photo = data.basics.photo || "";
  const hasPhoto = PHOTO_PATTERN.test(photo);
  thumbnail.hidden = !hasPhoto;
  if (hasPhoto) setPhoto(thumbnail, photo, data.basics.photo_focus);
  else {
    thumbnail.onload = null;
    thumbnail.removeAttribute("src");
    thumbnail.style.objectPosition = "50% 50%";
  }
  byId("photo-status").textContent = hasPhoto ? "已添加照片，页眉采用照片、姓名、联系方式三栏布局。" : "未添加照片，保持当前双栏页眉。";
}

async function api(path, payload) {
  const response = await fetch(path, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(payload) });
  if (!response.ok) {
    let message = "请求失败，请稍后重试。";
    try { message = (await response.json()).error || message; } catch (_) { /* Keep the readable fallback. */ }
    throw new Error(message);
  }
  return response;
}

async function validate(payload, { replaceEditor = false, syncForm = true, preserveOnFailure = false } = {}) {
  const previousValid = state.valid;
  const serial = ++state.serial;
  state.valid = false;
  updateButtons();
  status("正在检查配置…", true);
  try {
    const response = await api("/api/preview", payload);
    const result = await response.json();
    if (serial !== state.serial) return;
    state.resume = result.resume;
    state.canonical = result.yaml;
    state.valid = true;
    if (replaceEditor) editor.value = result.yaml;
    if (syncForm) fillForm(result.resume);
    render(result.resume);
    error();
    status("配置已通过检查，预览已更新。");
  } catch (failure) {
    if (serial !== state.serial) return;
    state.valid = preserveOnFailure && previousValid;
    error(failure.message);
    status(state.valid ? "本次更新未应用，原有配置仍可导出。" : "请修正配置后再导出。", true);
  } finally {
    if (serial === state.serial) updateButtons();
  }
}

function schedule(payload, options) {
  clearTimeout(state.timer);
  state.serial += 1;
  state.valid = false;
  updateButtons();
  status("等待内容检查…", true);
  state.timer = setTimeout(() => validate(payload, options), 350);
}

function download(contents, filename, type) {
  const blob = contents instanceof Blob ? contents : new Blob([contents], { type });
  const url = URL.createObjectURL(blob);
  const link = element("a");
  link.href = url;
  link.download = filename;
  document.body.append(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

form.addEventListener("submit", (event) => event.preventDefault());
form.addEventListener("input", (event) => {
  if (event.target.type === "file" || !event.target.name) return;
  if (!state.resume || state.busy) return;
  const data = JSON.parse(JSON.stringify(state.resume));
  for (const field of form.elements) {
    if (!field.name) continue;
    if (field.name === "profile") data.profile = field.value;
    else data.basics[field.name] = field.value;
  }
  schedule({ resume: data }, { replaceEditor: true, syncForm: false });
});
editor.addEventListener("input", () => schedule({ yaml: editor.value }, { replaceEditor: false, syncForm: true }));
editor.addEventListener("keydown", (event) => {
  if (event.key === "Tab") {
    event.preventDefault();
    editor.setRangeText("  ", editor.selectionStart, editor.selectionEnd, "end");
    editor.dispatchEvent(new Event("input"));
  }
  if ((event.ctrlKey || event.metaKey) && event.key === "Enter") { event.preventDefault(); byId("format-button").click(); }
});
byId("format-button").addEventListener("click", () => { clearTimeout(state.timer); validate({ yaml: editor.value }, { replaceEditor: true }); });
byId("preset").addEventListener("change", () => {
  const preset = state.presets.find((item) => item.id === byId("preset").value);
  if (!preset) return;
  clearTimeout(state.timer);
  editor.value = preset.yaml;
  validate({ resume: preset.resume }, { replaceEditor: true });
});
byId("import-button").addEventListener("click", () => byId("import-file").click());
byId("import-file").addEventListener("change", async (event) => {
  const file = event.target.files[0];
  if (!file) return;
  try {
    if (file.size > 6 * 1024 * 1024) throw new Error("配置文件不能超过 6 MB。");
    clearTimeout(state.timer);
    editor.value = await file.text();
    await validate({ yaml: editor.value }, { replaceEditor: false });
  } catch (failure) { error(failure.message); }
  event.target.value = "";
});
byId("save-config").addEventListener("click", () => {
  if (state.valid) download(state.canonical, "my_resume.yaml", "application/yaml;charset=utf-8");
});

byId("upload-photo").addEventListener("click", () => byId("photo-file").click());
byId("photo-file").addEventListener("change", async (event) => {
  const file = event.target.files[0];
  if (!file) return;
  const serial = state.serial;
  try {
    if (!state.valid || state.busy) throw new Error("请先完成当前内容的校验，再上传照片。");
    if (!["image/png", "image/jpeg", "image/webp"].includes(file.type)) throw new Error("请选择 PNG、JPEG 或 WebP 格式的照片。");
    if (file.size > 2 * 1024 * 1024) throw new Error("照片不能超过 2 MB，请压缩后再上传。");
    const uri = await new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result);
      reader.onerror = () => reject(new Error("照片读取失败，请重新选择。"));
      reader.onabort = () => reject(new Error("照片读取已取消。"));
      reader.readAsDataURL(file);
    });
    if (serial !== state.serial || !state.valid || state.busy) throw new Error("内容已发生变化，请等待校验完成后重新上传照片。");
    if (typeof uri !== "string" || !PHOTO_PATTERN.test(uri)) throw new Error("无法识别这张照片，请换一张图片。");
    const data = JSON.parse(JSON.stringify(state.resume));
    data.basics.photo = uri;
    data.basics.photo_focus = [0.5, 0.5];
    clearTimeout(state.timer);
    await validate({ resume: data }, { replaceEditor: true, preserveOnFailure: true });
  } catch (failure) { error(failure.message); }
  event.target.value = "";
});
byId("remove-photo").addEventListener("click", () => {
  if (!state.valid || state.busy || !state.resume?.basics.photo) return;
  const data = JSON.parse(JSON.stringify(state.resume));
  data.basics.photo = "";
  data.basics.photo_focus = [0.5, 0.5];
  clearTimeout(state.timer);
  validate({ resume: data }, { replaceEditor: true });
});

for (const button of document.querySelectorAll("[data-add]")) {
  button.addEventListener("click", () => {
    if (!state.valid || state.busy) return;
    const data = JSON.parse(JSON.stringify(state.resume));
    const key = button.dataset.add;
    if (key === "education") data.education.push({ date: "20XX.09 - 至今", school: "学校名称", major: "专业与学历", detail: "" });
    else if (key === "skills") data.skills.groups.push({ label: "技能分类", value: "填写实际使用的技术" });
    else if (key === "certificates") data.certificates.push({ name: "奖项或证书名称", result: "实际等级、成绩或年份" });
    else data[key].push({ name: key === "internship_experiences" ? "公司或团队名称" : "项目名称", date: "20XX.XX - 至今", role: "你的职责", description: "介绍服务对象、解决的问题和主要功能。", highlights: ["个人贡献：写清做了什么、如何验证以及实际结果。"], tech_stack: "", url: "" });
    clearTimeout(state.timer);
    validate({ resume: data }, { replaceEditor: true }).then(() => {
      const marker = key === "skills" ? "skills:" : key + ":";
      const index = editor.value.indexOf(marker);
      editor.focus();
      editor.setSelectionRange(Math.max(index, 0), Math.max(index, 0));
      editor.scrollTop = Math.max(0, editor.value.slice(0, Math.max(index, 0)).split("\n").length - 3) * 19;
    });
  });
}

for (const format of ["docx", "pdf"]) {
  byId("export-" + format).addEventListener("click", async () => {
    if (!state.valid || state.busy) return;
    const exportData = JSON.parse(JSON.stringify(state.resume));
    state.busy = true;
    updateButtons();
    const button = byId("export-" + format);
    const original = button.textContent;
    button.textContent = "正在生成…";
    status(format === "pdf" ? "正在本机转换 PDF，首次启动 Office 可能需要一些时间…" : "正在生成可编辑的 Word 文件…", true);
    error();
    try {
      const response = await api("/api/export?format=" + format, { resume: exportData });
      download(await response.blob(), exportData.basics.name.replace(/[<>:"/\\|?*\x00-\x1f]/g, "_") + "." + format);
      status("文件已生成，下载已交给浏览器。");
    } catch (failure) {
      error(failure.message);
      status("本次导出未完成；配置与预览仍保留。", true);
    } finally {
      state.busy = false;
      button.textContent = original;
      updateButtons();
    }
  });
}

async function init() {
  try {
    const response = await fetch("/api/examples");
    if (!response.ok) throw new Error((await response.json()).error || "无法加载模板。");
    const data = await response.json();
    state.presets = data.examples;
    const select = byId("preset");
    select.replaceChildren();
    for (const preset of state.presets) { const option = element("option", "", preset.name); option.value = preset.id; select.append(option); }
    const preset = state.presets.find((item) => item.id === "anon") || state.presets[0];
    select.value = preset.id;
    editor.value = preset.yaml;
    await validate({ resume: preset.resume }, { replaceEditor: true });
  } catch (failure) { error(failure.message); status("编辑器初始化失败，请确认本地服务仍在运行。", true); }
  updateButtons();
}

init();

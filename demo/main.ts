import { createEditor, sniffFormat, type RichEditor } from "../src/index";

const fileInput = document.getElementById("file") as HTMLInputElement;
const mount = document.getElementById("mount") as HTMLElement;
const info = document.getElementById("info") as HTMLElement;
const saveBtn = document.getElementById("save") as HTMLButtonElement;

let editor: RichEditor | null = null;
let currentName = "document";

fileInput.addEventListener("change", async () => {
  const f = fileInput.files?.[0];
  if (!f) return;
  currentName = f.name;
  const bytes = new Uint8Array(await f.arrayBuffer());
  editor?.destroy();
  mount.innerHTML = "";
  const fmt = sniffFormat(bytes);
  editor = createEditor(mount, bytes, {
    author: "Demo User",
    onChange: () => { info.textContent = `${fmt} — ${f.name} — edited`; },
  });
  info.textContent = `${fmt} — ${f.name} (${bytes.length} bytes)`;
  saveBtn.disabled = false;
});

saveBtn.addEventListener("click", async () => {
  if (!editor) return;
  const out = await editor.getBytes();
  const a = document.createElement("a");
  a.href = URL.createObjectURL(new Blob([out]));
  a.download = currentName;
  a.click();
});

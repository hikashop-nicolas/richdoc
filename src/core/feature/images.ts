// Image handling: click-to-select with a corner resize handle and a delete button (or the
// Delete/Backspace key), persisting the new size back into the run, plus inserting a new
// image from a file as an inline data-URL <img>. Extracted from the engine.
import { t } from "../i18n";
import { bytesToBase64 } from "../util";

export interface ImagesDeps {
  wrap: HTMLElement;
  scroll: HTMLElement;
  regions: HTMLElement[];
  mark: () => void;
  getActiveEl: () => HTMLElement;
  /** Notified when the selected image changes (null = deselected), for the layout toolbar. */
  onSelect?: (img: HTMLImageElement | null) => void;
}

export function setupImages(deps: ImagesDeps) {
  const { wrap, scroll, regions, mark, getActiveEl, onSelect } = deps;

  // Image select: click an image to select it, drag the corner handle to resize, and use
  // the delete button (top-right) or the Delete/Backspace key to remove it.
  let selImg: HTMLImageElement | null = null;
  const imgHandle = document.createElement("div");
  imgHandle.className = "docxedit-img-handle is-hidden";
  const imgDel = document.createElement("button");
  imgDel.type = "button";
  imgDel.className = "docxedit-img-del is-hidden";
  imgDel.textContent = "✕";
  imgDel.title = t("deleteImage");
  wrap.append(imgHandle, imgDel);
  const placeHandle = () => {
    if (!selImg || !wrap.contains(selImg)) {
      imgHandle.classList.add("is-hidden");
      imgDel.classList.add("is-hidden");
      return;
    }
    const ir = selImg.getBoundingClientRect();
    const wr = wrap.getBoundingClientRect();
    // positions are runtime-computed; they must follow the selected image
    imgHandle.style.left = `${ir.right - wr.left - 6}px`;
    imgHandle.style.top = `${ir.bottom - wr.top - 6}px`;
    imgHandle.classList.remove("is-hidden");
    imgDel.style.left = `${ir.right - wr.left - 11}px`;
    imgDel.style.top = `${ir.top - wr.top - 11}px`;
    imgDel.classList.remove("is-hidden");
  };
  const selectImg = (img: HTMLImageElement | null) => {
    if (selImg) selImg.classList.remove("sel");
    selImg = img;
    if (selImg) selImg.classList.add("sel");
    placeHandle();
    onSelect?.(selImg);
  };
  const deleteSelImg = () => {
    if (!selImg) return;
    selImg.remove();
    selImg = null;
    placeHandle();
    mark();
  };
  imgDel.addEventListener("mousedown", (e) => e.preventDefault());
  imgDel.addEventListener("click", (e) => {
    e.stopPropagation();
    deleteSelImg();
  });
  wrap.addEventListener("click", (e) => {
    if ((e.target as HTMLElement) === imgDel) return;
    const img = (e.target as HTMLElement).closest?.("img") as HTMLImageElement | null;
    selectImg(img && wrap.contains(img) ? img : null);
  });
  for (const r of regions)
    r.addEventListener("keydown", (e) => {
      if (selImg && (e.key === "Delete" || e.key === "Backspace")) {
        e.preventDefault();
        deleteSelImg();
      }
    });
  scroll.addEventListener("scroll", placeHandle);
  // Persist a resize: new images use the width/height attributes; existing ones carry the
  // original drawing in data-docx-xml, so update its extent (EMU) to the new size.
  const persistImgSize = (img: HTMLImageElement) => {
    const xml = img.getAttribute("data-docx-xml");
    const w = Number(img.getAttribute("width")) || 0;
    const h = Number(img.getAttribute("height")) || 0;
    if (!xml || !w) return;
    try {
      const frag = new DOMParser().parseFromString(xml, "application/xml");
      for (const tag of ["wp:extent", "a:ext"]) {
        const el = frag.getElementsByTagName(tag)[0];
        if (el) {
          el.setAttribute("cx", String(Math.round(w * 9525)));
          el.setAttribute("cy", String(Math.round(h * 9525)));
        }
      }
      img.setAttribute("data-docx-xml", new XMLSerializer().serializeToString(frag.documentElement!));
    } catch {
      /* leave as-is */
    }
  };
  imgHandle.addEventListener("mousedown", (e) => {
    e.preventDefault();
    const img = selImg;
    if (!img) return;
    const startX = e.clientX;
    const rect = img.getBoundingClientRect();
    const startW = rect.width;
    const ratio = rect.height / startW || 1;
    const onMove = (ev: MouseEvent) => {
      const w = Math.max(16, Math.round(startW + (ev.clientX - startX)));
      img.setAttribute("width", String(w));
      img.setAttribute("height", String(Math.round(w * ratio)));
      placeHandle();
    };
    const onUp = () => {
      document.removeEventListener("mousemove", onMove);
      document.removeEventListener("mouseup", onUp);
      persistImgSize(img);
      mark();
    };
    document.addEventListener("mousemove", onMove);
    document.addEventListener("mouseup", onUp);
  });

  const insertImage = () => {
    const sel = window.getSelection();
    const savedRange = sel && sel.rangeCount ? sel.getRangeAt(0).cloneRange() : null;
    const fileInput = document.createElement("input");
    fileInput.type = "file";
    fileInput.accept = "image/png,image/jpeg,image/gif,image/bmp,image/webp";
    fileInput.addEventListener("change", async () => {
      const file = fileInput.files && fileInput.files[0];
      if (!file) return;
      const buf = new Uint8Array(await file.arrayBuffer());
      const dataUrl = `data:${file.type};base64,${bytesToBase64(buf)}`;
      const probe = new Image();
      probe.onload = () => {
        const maxW = 600;
        let w = probe.naturalWidth || 200;
        let h = probe.naturalHeight || 200;
        if (w > maxW) {
          h = Math.round((h * maxW) / w);
          w = maxW;
        }
        const img = document.createElement("img");
        img.src = dataUrl;
        img.setAttribute("width", String(w));
        img.setAttribute("height", String(h));
        const range = savedRange ?? (() => {
          const r = document.createRange();
          r.selectNodeContents(getActiveEl());
          r.collapse(false);
          return r;
        })();
        range.collapse(false);
        range.insertNode(img);
        range.setStartAfter(img);
        range.collapse(true);
        const s2 = window.getSelection();
        if (s2) {
          s2.removeAllRanges();
          s2.addRange(range);
        }
        mark();
      };
      probe.src = dataUrl;
    });
    fileInput.click();
  };

  return { insertImage, repositionHandles: placeHandle };
}

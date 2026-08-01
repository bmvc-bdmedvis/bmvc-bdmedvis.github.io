import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const ROOT = "/Users/huyang/Workspace/bmvc-bdmedvis.github.io";
const TMP = `${ROOT}/poster/.v2-build/organizer-revision`;
const SOURCE = `${TMP}/template-starter.pptx`;
const FINAL_PPTX = `${ROOT}/poster/BD-MedVis-2026-A4-Poster-Editable-v4.pptx`;
const FINAL_PNG = `${ROOT}/poster/BD-MedVis-2026-A4-Poster-v4.png`;

const organizers = [
  { name: "Dr Yang Hu", file: "assets/organizers/yang-hu.jpg" },
  { name: "Dr Tianyang Zhang", file: "assets/organizers/tianyang-zhang.jpg" },
  { name: "Dr He Zhao", file: "assets/organizers/he-zhao.jpg" },
  { name: "Dr Jian Chen", file: "assets/organizers/jian-chen.jpg" },
  { name: "Dr Dan Dai", file: "assets/organizers/dan-dai.jpg" },
  { name: "Dr Yakun Ju", file: "assets/organizers/yakun-ju.jpg" },
  { name: "Dr Le Zhang", file: "assets/organizers/le-zhang.jpg" },
  { name: "Dr Shangqi Gao", file: "assets/organizers/shangqi-gao.jpg" },
  { name: "Dr Zheheng Jiang", file: "assets/organizers/zheheng-jiang.jpg" },
];

async function writeBlob(filePath, blob) {
  await fs.writeFile(filePath, new Uint8Array(await blob.arrayBuffer()));
}

async function readBytes(relPath) {
  return new Uint8Array(await fs.readFile(path.join(ROOT, relPath)));
}

async function main() {
  const presentation = await PresentationFile.importPptx(await FileBlob.load(SOURCE));
  const snapshot = await presentation.inspect({
    kind: "slide,textbox,shape,image,notes",
    include: "id,slide,name,bbox,text,textPreview,alt",
    maxChars: 120000,
  });
  const records = snapshot.ndjson.trim().split("\n").filter(Boolean).map((line) => JSON.parse(line));
  const recordByName = (name) => records.find((record) => record.name === name);
  const slideRecord = records.find((record) => record.kind === "slide" && record.slide === 1);
  if (!slideRecord) throw new Error("Could not resolve the poster slide.");
  const slide = presentation.resolve(slideRecord.id);

  await writeBlob(`${TMP}/before-organizers.png`, await presentation.export({ slide, format: "png", scale: 2 }));
  await fs.writeFile(`${TMP}/before-organizers.layout.json`, await (await slide.export({ format: "layout" })).text());

  const columnXs = [40, 278, 516];
  const rowYs = [846, 876, 906];

  for (let i = 0; i < organizers.length; i++) {
    const row = Math.floor(i / 3);
    const col = i % 3;
    const x = columnXs[col];
    const y = rowYs[row];

    const nameRecord = recordByName(`organizer-${i + 1}-name`);
    const affiliationRecord = recordByName(`organizer-${i + 1}-affiliation`);
    if (!nameRecord || !affiliationRecord) throw new Error(`Could not resolve organizer ${i + 1} text boxes.`);

    presentation.resolve(nameRecord.id).position = {
      left: x + 36,
      top: y + 1,
      width: 174,
      height: 13,
    };
    presentation.resolve(affiliationRecord.id).position = {
      left: x + 36,
      top: y + 15,
      width: 174,
      height: 12,
    };

    const portrait = slide.images.add({
      blob: await readBytes(organizers[i].file),
      contentType: "image/jpeg",
      alt: `${organizers[i].name}, workshop organizer`,
      fit: "cover",
      position: { left: x, top: y, width: 28, height: 28 },
      geometry: "ellipse",
    });
    portrait.name = `organizer-${i + 1}-photo`;
  }

  const notesRecord = records.find((record) => record.kind === "notes" && record.slide === 1);
  const existingNotes = notesRecord?.text ?? notesRecord?.textPreview ?? "";
  if (existingNotes.includes("[/Sources]") && !existingNotes.includes("assets/organizers")) {
    slide.speakerNotes.textFrame.setText(
      existingNotes.replace(
        "[/Sources]",
        `- ${ROOT}/assets/organizers (nine organizer portraits; accessed 2026-08-01)\n[/Sources]`
      )
    );
  }

  await writeBlob(`${TMP}/after-organizers.png`, await presentation.export({ slide, format: "png", scale: 2 }));
  await fs.mkdir(`${TMP}/final-layout`, { recursive: true });
  await fs.writeFile(`${TMP}/final-layout/final-slide-01.layout.json`, await (await slide.export({ format: "layout" })).text());
  await writeBlob(FINAL_PNG, await presentation.export({ slide, format: "png", scale: 3.125 }));

  const pptx = await PresentationFile.exportPptx(presentation);
  await pptx.save(FINAL_PPTX);

  const after = await presentation.inspect({
    kind: "slide,textbox,image,notes",
    search: "organizer-",
    maxChars: 30000,
  });
  await fs.writeFile(`${TMP}/after-inspect.ndjson`, after.ndjson);
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});

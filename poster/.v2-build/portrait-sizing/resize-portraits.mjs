import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const ROOT = "/Users/huyang/Workspace/bmvc-bdmedvis.github.io";
const TMP = `${ROOT}/poster/.v2-build/portrait-sizing`;
const SOURCE = `${TMP}/template-starter.pptx`;
const FINAL_PPTX = `${ROOT}/poster/BD-MedVis-2026-A4-Poster-Editable-v5.pptx`;
const FINAL_PNG = `${ROOT}/poster/BD-MedVis-2026-A4-Poster-v5.png`;

async function writeBlob(filePath, blob) {
  await fs.writeFile(filePath, new Uint8Array(await blob.arrayBuffer()));
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

  await writeBlob(`${TMP}/before-sizing.png`, await presentation.export({ slide, format: "png", scale: 2 }));
  await fs.writeFile(`${TMP}/before-sizing.layout.json`, await (await slide.export({ format: "layout" })).text());

  const columnXs = [40, 278, 516];
  const rowYs = [846, 876, 906];
  for (let i = 0; i < 9; i++) {
    const portraitRecord = recordByName(`organizer-${i + 1}-photo`);
    if (!portraitRecord) throw new Error(`Could not resolve organizer portrait ${i + 1}.`);
    const row = Math.floor(i / 3);
    const col = i % 3;
    presentation.resolve(portraitRecord.id).frame = {
      left: columnXs[col] + 3,
      top: rowYs[row] + 3,
      width: 22,
      height: 22,
    };
  }

  await writeBlob(`${TMP}/after-sizing.png`, await presentation.export({ slide, format: "png", scale: 2 }));
  await fs.mkdir(`${TMP}/final-layout`, { recursive: true });
  await fs.writeFile(`${TMP}/final-layout/final-slide-01.layout.json`, await (await slide.export({ format: "layout" })).text());
  await writeBlob(FINAL_PNG, await presentation.export({ slide, format: "png", scale: 3.125 }));

  const pptx = await PresentationFile.exportPptx(presentation);
  await pptx.save(FINAL_PPTX);

  const after = await presentation.inspect({
    kind: "slide,image",
    search: "organizer-",
    maxChars: 30000,
  });
  await fs.writeFile(`${TMP}/after-inspect.ndjson`, after.ndjson);
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});

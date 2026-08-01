import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const ROOT = "/Users/huyang/Workspace/bmvc-bdmedvis.github.io";
const TMP = `${ROOT}/poster/.v2-build/hero-revision`;
const SOURCE = `${TMP}/template-starter.pptx`;
const FINAL_PPTX = `${ROOT}/poster/BD-MedVis-2026-A4-Poster-Editable-v3.pptx`;
const FINAL_PNG = `${ROOT}/poster/BD-MedVis-2026-A4-Poster-v3.png`;

const C = {
  ink: "#102329",
  muted: "#607176",
  teal: "#005D66",
};

async function writeBlob(filePath, blob) {
  await fs.writeFile(filePath, new Uint8Array(await blob.arrayBuffer()));
}

async function main() {
  const presentation = await PresentationFile.importPptx(await FileBlob.load(SOURCE));
  const snapshot = await presentation.inspect({
    kind: "slide,textbox,shape",
    include: "id,slide,name,bbox,text,textPreview",
    maxChars: 40000,
  });
  const records = snapshot.ndjson.trim().split("\n").filter(Boolean).map((line) => JSON.parse(line));
  const recordByName = (name) => records.find((record) => record.name === name);
  const slideRecord = records.find((record) => record.kind === "slide" && record.slide === 1);
  if (!slideRecord) throw new Error("Could not resolve the poster slide.");
  const slide = presentation.resolve(slideRecord.id);

  await writeBlob(`${TMP}/before-hero.png`, await presentation.export({ slide, format: "png", scale: 2 }));
  await fs.writeFile(`${TMP}/before-hero.layout.json`, await (await slide.export({ format: "layout" })).text());

  const kickerRecord = recordByName("hero-kicker");
  const titleRecord = recordByName("hero-title");
  const copyRecord = recordByName("hero-copy");
  const eventDateLabelRecord = recordByName("event-label-0");
  if (!kickerRecord || !titleRecord || !copyRecord || !eventDateLabelRecord) throw new Error("Could not resolve the inherited poster text boxes.");

  const kicker = presentation.resolve(kickerRecord.id);
  kicker.text = "BMVC 2026 WORKSHOP";
  kicker.position = { left: 38, top: 106, width: 220, height: 15 };
  kicker.text.style = {
    typeface: "Aptos",
    fontSize: 8.8,
    bold: true,
    color: C.teal,
    alignment: "left",
    verticalAlignment: "top",
    autoFit: "shrinkText",
    lineSpacing: 0.95,
    insets: { top: 0, right: 0, bottom: 0, left: 0 },
  };

  const title = presentation.resolve(titleRecord.id);
  title.text = "Biomarker-Driven\nMedical Vision\nIntelligence";
  title.position = { left: 38, top: 124, width: 390, height: 116 };
  title.text.style = {
    typeface: "Aptos Display",
    fontSize: 34,
    bold: true,
    color: C.ink,
    alignment: "left",
    verticalAlignment: "top",
    autoFit: "shrinkText",
    lineSpacing: 0.88,
    insets: { top: 0, right: 0, bottom: 0, left: 0 },
  };

  const copy = presentation.resolve(copyRecord.id);
  copy.position = { left: 40, top: 246, width: 388, height: 82 };
  copy.text = [
    {
      runs: [
        {
          run: "From pixels to clinically credible biomarkers.",
          textStyle: {
            bold: true,
            fontSize: "13.2px",
            typeface: "Aptos Display",
            color: C.teal,
          },
        },
      ],
      spaceAfter: 4,
    },
    {
      runs: [
        {
          run: "BD-MedVis brings together computer vision, medical imaging and clinical AI to create image-derived biomarkers that are biologically grounded, reproducible and clinically meaningful. Join researchers across radiology, pathology and multimodal medicine to advance trustworthy methods for precision care.",
          textStyle: {
            fontSize: "9.2px",
            typeface: "Aptos",
            color: C.muted,
          },
        },
      ],
    },
  ];
  copy.text.style = {
    typeface: "Aptos",
    fontSize: 9.2,
    color: C.muted,
    alignment: "left",
    verticalAlignment: "top",
    autoFit: "shrinkText",
    lineSpacing: 1.0,
    insets: { top: 0, right: 0, bottom: 0, left: 0 },
  };

  presentation.resolve(eventDateLabelRecord.id).text = "WORKSHOP DATE";

  await writeBlob(`${TMP}/after-hero.png`, await presentation.export({ slide, format: "png", scale: 2 }));
  await fs.writeFile(`${TMP}/final-layout/final-slide-01.layout.json`, await (await slide.export({ format: "layout" })).text());
  await writeBlob(FINAL_PNG, await presentation.export({ slide, format: "png", scale: 3.125 }));

  const pptx = await PresentationFile.exportPptx(presentation);
  await pptx.save(FINAL_PPTX);

  const after = await presentation.inspect({
    kind: "slide,textbox,shape,notes",
    search: "Biomarker-Driven|From pixels|BD-MedVis brings",
    maxChars: 8000,
  });
  await fs.writeFile(`${TMP}/after-inspect.ndjson`, after.ndjson);
}

await fs.mkdir(`${TMP}/final-layout`, { recursive: true });
main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});

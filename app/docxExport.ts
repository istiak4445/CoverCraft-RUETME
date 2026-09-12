import {
  AlignmentType,
  BorderStyle,
  Document,
  ImageRun,
  Packer,
  PageBreak,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from "docx";

export type DocxCover = {
  type: string;
  courseNo: string;
  courseTitle: string;
  itemNo: string;
  itemName: string;
  name: string;
  roll: string;
  section: string;
  session: string;
  date: string;
  teacher: string;
  designation: string;
};

const PAGE_WIDTH = 11906;
const PAGE_HEIGHT = 16838;
const MARGIN = 1134;
const CONTENT_WIDTH = PAGE_WIDTH - MARGIN * 2;
const blackBorder = { style: BorderStyle.SINGLE, size: 8, color: "111111" };

function detail(label: string, value: string) {
  return new Paragraph({
    spacing: { after: 210 },
    children: [
      new TextRun({
        text: `${label}: `,
        bold: true,
        size: 30,
        font: "Times New Roman",
      }),
      new TextRun({ text: value, size: 30, font: "Times New Roman" }),
    ],
  });
}

function lines(values: string[]) {
  return new Paragraph({
    spacing: { line: 300 },
    children: values.flatMap((value, index) => [
      new TextRun({ text: value, size: 27, font: "Times New Roman" }),
      ...(index < values.length - 1 ? [new TextRun({ break: 1 })] : []),
    ]),
  });
}

function coverChildren(cover: DocxCover, logo: Uint8Array, addBreak: boolean) {
  const children = [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 100 },
      children: [
        new TextRun({
          text: "Rajshahi University of Engineering & Technology",
          bold: true,
          size: 42,
          font: "Times New Roman",
        }),
      ],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 380 },
      children: [
        new TextRun({
          text: "Department of Mechanical Engineering",
          bold: true,
          size: 34,
          font: "Times New Roman",
        }),
      ],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 950 },
      children: [
        new ImageRun({
          data: logo,
          type: "png",
          transformation: { width: 145, height: 145 },
          altText: {
            title: "RUET logo",
            description:
              "Rajshahi University of Engineering and Technology logo",
            name: "RUET logo",
          },
        }),
      ],
    }),
    detail("Course No", cover.courseNo),
    detail("Course Title", cover.courseTitle),
    detail(`${cover.type} No`, cover.itemNo),
    detail(`${cover.type} Name`, cover.itemName),
    new Paragraph({ spacing: { after: 260 } }),
    new Table({
      width: { size: CONTENT_WIDTH, type: WidthType.DXA },
      columnWidths: [CONTENT_WIDTH / 2, CONTENT_WIDTH / 2],
      borders: {
        top: blackBorder,
        bottom: blackBorder,
        left: blackBorder,
        right: blackBorder,
        insideHorizontal: blackBorder,
        insideVertical: blackBorder,
      },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              width: { size: CONTENT_WIDTH / 2, type: WidthType.DXA },
              margins: { top: 100, bottom: 100, left: 140, right: 140 },
              children: [
                new Paragraph({
                  spacing: { after: 320 },
                  children: [
                    new TextRun({
                      text: "Submitted By",
                      size: 29,
                      font: "Times New Roman",
                    }),
                  ],
                }),
                lines([
                  `Name: ${cover.name}`,
                  `Roll: ${cover.roll}`,
                  `Section: ${cover.section}`,
                  `Session: ${cover.session}`,
                  `Date of submission: ${cover.date}`,
                ]),
              ],
            }),
            new TableCell({
              width: { size: CONTENT_WIDTH / 2, type: WidthType.DXA },
              margins: { top: 100, bottom: 100, left: 140, right: 140 },
              children: [
                new Paragraph({
                  spacing: { after: 320 },
                  children: [
                    new TextRun({
                      text: "Submitted To",
                      size: 29,
                      font: "Times New Roman",
                    }),
                  ],
                }),
                lines([
                  cover.teacher,
                  cover.designation,
                  "Department of Mechanical Engineering,",
                  "RUET",
                ]),
              ],
            }),
          ],
        }),
      ],
    }),
  ];
  if (addBreak) children.push(new Paragraph({ children: [new PageBreak()] }));
  return children;
}

export async function exportCoversToDocx(covers: DocxCover[]) {
  const logoResponse = await fetch("/ruet-logo.png");
  if (!logoResponse.ok) throw new Error("Could not load RUET logo");
  const logo = new Uint8Array(await logoResponse.arrayBuffer());
  const children = covers.flatMap((cover, index) =>
    coverChildren(cover, logo, index < covers.length - 1),
  );
  const document = new Document({
    creator: "CoverCraft RUET Mechanical",
    title:
      covers.length === 1
        ? `${covers[0].type} Cover`
        : "RUET Mechanical Cover Pages",
    sections: [
      {
        properties: {
          page: {
            size: { width: PAGE_WIDTH, height: PAGE_HEIGHT },
            margin: {
              top: MARGIN,
              right: MARGIN,
              bottom: MARGIN,
              left: MARGIN,
            },
          },
        },
        children,
      },
    ],
  });
  const blob = await Packer.toBlob(document);
  const url = URL.createObjectURL(blob);
  const anchor = documentElement("a");
  anchor.href = url;
  anchor.download =
    covers.length === 1
      ? `${covers[0].courseNo.replaceAll(" ", "-")}-${covers[0].type.replaceAll(" ", "-")}.docx`
      : "RUET-ME-CoverCraft-Batch.docx";
  anchor.click();
  URL.revokeObjectURL(url);
}

function documentElement(tag: "a") {
  return window.document.createElement(tag);
}

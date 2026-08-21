import React, { useEffect, useState } from "react";
import { createRoot } from "react-dom/client";

import {
  BrowserRouter,
  Routes,
  Route,
  Link,
  useParams,
} from "react-router-dom";

import {
  Menu,
  X,
  Search,
  Upload,
  Merge,
  Scissors,
  Minimize2,
  FileText,
  Image as ImageIcon,
  RotateCw,
  LockKeyhole,
  UnlockKeyhole,
  Hash,
  Stamp,
  Trash2,
  ArrowDownToLine,
  ChevronRight,
  Check,
  Zap,
  Globe,
  Shield,
  PenTool,
  ListOrdered,
  Copy,
  Presentation,
  Table2,
  Code2,
  FileSpreadsheet,
} from "lucide-react";

import { PDFDocument, degrees, rgb, StandardFonts } from "pdf-lib";
import * as pdfjsLib from "pdfjs-dist";
import jsPDF from "jspdf";
import { Document, Packer, Paragraph, TextRun } from "docx";
import PptxGenJS from "pptxgenjs";
import * as XLSX from "xlsx";
import JSZip from "jszip";
import html2canvas from "html2canvas";

import { encryptPDF } from "@pdfsmaller/pdf-encrypt";
import { decryptPDF, isEncrypted } from "@pdfsmaller/pdf-decrypt";

import "./styles.css";

/* =========================================================
   PDF.JS WORKER
========================================================= */

pdfjsLib.GlobalWorkerOptions.workerSrc = new URL(
  "pdfjs-dist/build/pdf.worker.min.mjs",
  import.meta.url
).toString();

/* =========================================================
   TOOLS
========================================================= */

const tools = [
  {
    id: "merge-pdf",
    name: "Merge PDF",
    desc: "Combine PDFs in the order you want.",
    icon: Merge,
    cat: "Organize PDF",
  },
  {
    id: "split-pdf",
    name: "Split PDF",
    desc: "Separate PDF pages into individual files.",
    icon: Scissors,
    cat: "Organize PDF",
  },
  {
    id: "compress-pdf",
    name: "Compress PDF",
    desc: "Reduce PDF file size for easy sharing.",
    icon: Minimize2,
    cat: "Optimize PDF",
  },

  {
    id: "jpg-to-pdf",
    name: "JPG to PDF",
    desc: "Convert images into a PDF document.",
    icon: ImageIcon,
    cat: "Convert PDF",
  },
  {
    id: "pdf-to-jpg",
    name: "PDF to JPG",
    desc: "Convert PDF pages into images.",
    icon: ImageIcon,
    cat: "Convert PDF",
  },

  {
    id: "word-to-pdf",
    name: "WORD to PDF",
    desc: "Convert Word documents into PDF files.",
    icon: FileText,
    cat: "Convert PDF",
  },
  {
    id: "pdf-to-word",
    name: "PDF to WORD",
    desc: "Convert PDF text into an editable Word document.",
    icon: FileText,
    cat: "Convert PDF",
  },

  {
    id: "powerpoint-to-pdf",
    name: "POWERPOINT to PDF",
    desc: "Convert PowerPoint presentations into PDF.",
    icon: Presentation,
    cat: "Convert PDF",
  },
  {
    id: "pdf-to-powerpoint",
    name: "PDF to POWERPOINT",
    desc: "Convert PDF pages into a PowerPoint presentation.",
    icon: Presentation,
    cat: "Convert PDF",
  },

  {
    id: "excel-to-pdf",
    name: "EXCEL to PDF",
    desc: "Convert Excel spreadsheets into PDF files.",
    icon: FileSpreadsheet,
    cat: "Convert PDF",
  },
  {
    id: "pdf-to-excel",
    name: "PDF to EXCEL",
    desc: "Extract PDF text into an Excel spreadsheet.",
    icon: Table2,
    cat: "Convert PDF",
  },

  {
    id: "html-to-pdf",
    name: "HTML to PDF",
    desc: "Convert an HTML file into a PDF.",
    icon: Code2,
    cat: "Convert PDF",
  },
  {
    id: "pdf-to-pdfa",
    name: "PDF to PDF/A",
    desc: "Create a long-term archival PDF copy.",
    icon: FileText,
    cat: "Convert PDF",
  },

  {
    id: "rotate-pdf",
    name: "Rotate PDF",
    desc: "Rotate all PDF pages.",
    icon: RotateCw,
    cat: "Edit PDF",
  },
  {
    id: "watermark-pdf",
    name: "Watermark PDF",
    desc: "Add a text watermark to every page.",
    icon: Stamp,
    cat: "Edit PDF",
  },
  {
    id: "page-numbers",
    name: "Page Numbers",
    desc: "Add page numbers to your PDF.",
    icon: Hash,
    cat: "Edit PDF",
  },

  {
    id: "remove-pages",
    name: "Remove Pages",
    desc: "Remove selected pages from a PDF.",
    icon: Trash2,
    cat: "Organize PDF",
  },
  {
    id: "extract-pages",
    name: "Extract Pages",
    desc: "Extract selected pages into a new PDF.",
    icon: Copy,
    cat: "Organize PDF",
  },
  {
    id: "reorder-pages",
    name: "Reorder Pages",
    desc: "Change the order of PDF pages.",
    icon: ListOrdered,
    cat: "Organize PDF",
  },

  {
    id: "sign-pdf",
    name: "Sign PDF",
    desc: "Add your signature to a PDF.",
    icon: PenTool,
    cat: "Edit PDF",
  },
  {
    id: "protect-pdf",
    name: "Protect PDF",
    desc: "Protect your PDF with a password.",
    icon: LockKeyhole,
    cat: "PDF Security",
  },
  {
    id: "unlock-pdf",
    name: "Unlock PDF",
    desc: "Unlock a PDF using its valid password.",
    icon: UnlockKeyhole,
    cat: "PDF Security",
  },
];
const seoData = {
  "merge-pdf": {
    title: "Merge PDF Online Free — Combine PDF Files | DesiPDF",
    description:
      "Merge multiple PDF files online for free with DesiPDF. Combine PDF files in your preferred order quickly and securely in your browser.",
  },

  "split-pdf": {
    title: "Split PDF Online Free — Split PDF Pages | DesiPDF",
    description:
      "Split PDF files online for free with DesiPDF. Separate PDF pages into individual files quickly and easily.",
  },

  "compress-pdf": {
    title: "Compress PDF Online Free — Reduce PDF Size | DesiPDF",
    description:
      "Compress PDF files online for free with DesiPDF. Reduce PDF file size for easier sharing and storage.",
  },

  "jpg-to-pdf": {
    title: "JPG to PDF Converter Online Free | DesiPDF",
    description:
      "Convert JPG and PNG images to PDF online for free with DesiPDF. Create PDF documents from your images directly in your browser.",
  },

  "pdf-to-jpg": {
    title: "PDF to JPG Converter Online Free | DesiPDF",
    description:
      "Convert PDF pages to JPG images online for free with DesiPDF. Turn PDF pages into high-quality JPG images directly in your browser.",
  },

  "pdf-to-word": {
    title: "PDF to Word Converter Online Free | DesiPDF",
    description:
      "Convert PDF files to editable Word documents online for free with DesiPDF. Fast browser-based PDF to Word conversion.",
  },

  "pdf-to-powerpoint": {
    title: "PDF to PowerPoint Converter Online Free | DesiPDF",
    description:
      "Convert PDF pages to PowerPoint presentations online for free with DesiPDF. Create PPTX files directly in your browser.",
  },

  "rotate-pdf": {
    title: "Rotate PDF Online Free — Rotate PDF Pages | DesiPDF",
    description:
      "Rotate PDF pages online for free with DesiPDF. Quickly rotate your PDF pages and download the updated document.",
  },

  "watermark-pdf": {
    title: "Add Watermark to PDF Online Free | DesiPDF",
    description:
      "Add a text watermark to your PDF online for free with DesiPDF. Protect and brand your PDF documents easily.",
  },

  "page-numbers": {
    title: "Add Page Numbers to PDF Online Free | DesiPDF",
    description:
      "Add page numbers to PDF files online for free with DesiPDF. Number your PDF pages quickly and easily.",
  },

  "remove-pages": {
    title: "Remove Pages from PDF Online Free | DesiPDF",
    description:
      "Remove unwanted pages from PDF files online for free with DesiPDF. Select and remove PDF pages easily.",
  },

  "extract-pages": {
    title: "Extract Pages from PDF Online Free | DesiPDF",
    description:
      "Extract selected pages from PDF files online for free with DesiPDF. Create a new PDF from the pages you choose.",
  },

  "reorder-pages": {
    title: "Reorder PDF Pages Online Free | DesiPDF",
    description:
      "Reorder PDF pages online for free with DesiPDF. Change the order of your PDF pages and download the new document.",
  },

  "sign-pdf": {
    title: "Sign PDF Online Free — Add Signature to PDF | DesiPDF",
    description:
      "Sign PDF documents online for free with DesiPDF. Add a typed signature to your PDF directly in your browser.",
  },

  "protect-pdf": {
    title: "Protect PDF with Password Online Free | DesiPDF",
    description:
      "Protect PDF files with a password online using DesiPDF. Add PDF security and control document permissions.",
  },

  "unlock-pdf": {
    title: "Unlock PDF Online Free — Remove PDF Password | DesiPDF",
    description:
      "Unlock password-protected PDF files online with DesiPDF when you have the valid PDF password.",
  },
};
/* =========================================================
   HEADER
========================================================= */

function Header() {
  const [open, setOpen] = useState(false);

  return (
    <header className="header">
      <div className="nav container">

        <Link to="/" className="brand">
          <span className="brandMark">D</span>
          <span>
            Desi<span>PDF</span>
          </span>
        </Link>

        <nav className={open ? "mobileOpen" : ""}>
          <Link to="/" onClick={() => setOpen(false)}>
            Home
          </Link>

          <Link to="/tools" onClick={() => setOpen(false)}>
            All Tools
          </Link>

          <a href="/#how" onClick={() => setOpen(false)}>
            How it works
          </a>

          <a href="/#about" onClick={() => setOpen(false)}>
            About
          </a>
        </nav>

        <div className="navActions">
          <Link className="login" to="/tools">
            Try Tools
          </Link>

          <button
            className="menuBtn"
            onClick={() => setOpen(!open)}
          >
            {open ? <X /> : <Menu />}
          </button>
        </div>

      </div>
    </header>
  );
}

/* =========================================================
   HOME
========================================================= */

function Home() {
  const [q, setQ] = useState("");

  const filtered = tools.filter((t) =>
    `${t.name} ${t.desc} ${t.cat}`
      .toLowerCase()
      .includes(q.toLowerCase())
  );

  return (
    <>
      <Header />

      <main>

        <section className="hero">
          <div className="container heroInner">

            <div className="pill">
              <Zap size={15} />
              Simple. Fast. Free.
            </div>

            <h1>
              Every PDF tool you need.
              <br />
              <em>All in one place.</em>
            </h1>

            <p>
              Merge, split, compress, convert,
              edit and protect your PDFs.
            </p>

            <div className="heroSearch">
              <Search size={20} />

              <input
                value={q}
                onChange={(e) => setQ(e.target.value)}
                placeholder="Search for a PDF tool..."
              />
            </div>

          </div>
        </section>

        <section id="tools" className="section container">

          <div className="sectionHead">

            <div>
              <span className="eyebrow">
                DESIPDF TOOLS
              </span>

              <h2>
                Everything for your PDFs
              </h2>
            </div>

            <span className="count">
              {filtered.length} tools
            </span>

          </div>

          <div className="categories">
            {[
              "All",
              "Organize PDF",
              "Optimize PDF",
              "Convert PDF",
              "Edit PDF",
              "PDF Security",
            ].map((c) => (
              <a key={c} href="#tools">
                {c}
              </a>
            ))}
          </div>

          <div className="toolGrid">

            {filtered.map((tool) => (
              <ToolCard
                key={tool.id}
                tool={tool}
              />
            ))}

          </div>

        </section>

        <section id="how" className="how">

          <div className="container">

            <span className="eyebrow">
              HOW IT WORKS
            </span>

            <h2>
              PDF editing made simple
            </h2>

            <div className="steps">

              {[
                [
                  "01",
                  "Choose a tool",
                  "Pick the PDF task you want.",
                ],
                [
                  "02",
                  "Upload your file",
                  "Choose your PDF, Word, Excel or other file.",
                ],
                [
                  "03",
                  "Process & download",
                  "DesiPDF processes your file.",
                ],
              ].map((s) => (
                <div className="step" key={s[0]}>
                  <b>{s[0]}</b>
                  <h3>{s[1]}</h3>
                  <p>{s[2]}</p>
                </div>
              ))}

            </div>

          </div>

        </section>

        <section id="about" className="trust">

          <div className="container trustGrid">

            <div>
              <Shield size={30} />

              <h3>
                Privacy first
              </h3>

              <p>
                Your files are processed directly
                in your browser wherever possible.
              </p>
            </div>

            <div>
              <Zap size={30} />

              <h3>
                Fast workflow
              </h3>

              <p>
                Simple tools with no unnecessary steps.
              </p>
            </div>

            <div>
              <Globe size={30} />

              <h3>
                Works anywhere
              </h3>

              <p>
                Use DesiPDF on desktop, tablet or mobile.
              </p>
            </div>

          </div>

        </section>

      </main>

      <Footer />
    </>
  );
}

/* =========================================================
   TOOL CARD
========================================================= */

function ToolCard({ tool }) {

  const Icon = tool.icon;

  return (
    <Link
      className="toolCard"
      to={`/tool/${tool.id}`}
    >

      <div className="iconBox">
        <Icon size={24} />
      </div>

      <div className="toolText">

        <h3>
          {tool.name}
        </h3>

        <p>
          {tool.desc}
        </p>

        <span>
          {tool.cat}
        </span>

      </div>

      <ChevronRight
        className="arrow"
        size={20}
      />

    </Link>
  );
}

/* =========================================================
   FOOTER
========================================================= */

function Footer() {

  return (
    <footer>

      <div className="container footerGrid">

        <div>

          <Link to="/" className="brand">

            <span className="brandMark">
              D
            </span>

            <span>
              Desi<span>PDF</span>
            </span>

          </Link>

          <p>
            Your simple, modern PDF toolkit.
          </p>

        </div>

        <div>

          <h4>
            Tools
          </h4>

          <Link to="/tools">
            All PDF Tools
          </Link>

          <Link to="/tool/merge-pdf">
            Merge PDF
          </Link>

          <Link to="/tool/compress-pdf">
            Compress PDF
          </Link>

          <Link to="/tool/pdf-to-powerpoint">
            PDF to PowerPoint
          </Link>

          <Link to="/tool/pdf-to-excel">
            PDF to Excel
          </Link>

          <Link to="/tool/protect-pdf">
            Protect PDF
          </Link>

        </div>

        <div>

          <h4>
            Company
          </h4>

          <a href="/#about">
            About
          </a>

          <a href="/#how">
            How it works
          </a>

        </div>

      </div>

      <div className="container copyright">
        © 2026 DesiPDF. Built for the web.
      </div>

    </footer>
  );
}

/* =========================================================
   FILE DROP
========================================================= */

function FileDrop({
  multiple = false,
  accept = ".pdf",
  onFiles,
}) {

  const [drag, setDrag] =
    useState(false);

  return (
    <div
      className={`drop ${drag ? "drag" : ""}`}

      onDragOver={(e) => {
        e.preventDefault();
        setDrag(true);
      }}

      onDragLeave={() => {
        setDrag(false);
      }}

      onDrop={(e) => {

        e.preventDefault();
        setDrag(false);

        const selected =
          [...e.dataTransfer.files];

        onFiles(
          multiple
            ? selected
            : selected.slice(0, 1)
        );

      }}
    >

      <Upload size={38} />

      <h3>
        Drop your file
        {multiple ? "s" : ""}
        {" "}here
      </h3>

      <p>
        or click to browse
      </p>

      <label className="browse">

        Choose file
        {multiple ? "s" : ""}

        <input
          type="file"
          multiple={multiple}
          accept={accept}

          onChange={(e) => {

            const selected =
              [...e.target.files];

            onFiles(
              multiple
                ? selected
                : selected.slice(0, 1)
            );

          }}
        />

      </label>

    </div>
  );
}

/* =========================================================
   HELPERS
========================================================= */

function fileData(file) {

  return new Promise(
    (resolve, reject) => {

      const reader =
        new FileReader();

      reader.onload =
        () => resolve(
          reader.result
        );

      reader.onerror =
        reject;

      reader.readAsDataURL(file);

    }
  );
}

function download(blob, name) {

  const url =
    URL.createObjectURL(blob);

  const a =
    document.createElement("a");

  a.href = url;
  a.download = name;

  document.body.appendChild(a);

  a.click();

  a.remove();

  setTimeout(() => {
    URL.revokeObjectURL(url);
  }, 1500);
}

function sleep(ms) {

  return new Promise(
    (resolve) =>
      setTimeout(
        resolve,
        ms
      )
  );
}

/* =========================================================
   MERGE PDF
========================================================= */

async function mergePDF(files) {

  if (files.length < 2) {
    throw new Error(
      "Please select at least two PDF files."
    );
  }

  const output =
    await PDFDocument.create();

  for (const file of files) {

    const source =
      await PDFDocument.load(
        await file.arrayBuffer()
      );

    const pages =
      await output.copyPages(
        source,
        source.getPageIndices()
      );

    pages.forEach((page) => {
      output.addPage(page);
    });
  }

  return new Blob(
    [await output.save()],
    {
      type: "application/pdf",
    }
  );
}

/* =========================================================
   SPLIT PDF
========================================================= */

async function splitPDF(file) {

  const source =
    await PDFDocument.load(
      await file.arrayBuffer()
    );

  const result = [];

  for (
    let i = 0;
    i < source.getPageCount();
    i++
  ) {

    const output =
      await PDFDocument.create();

    const [page] =
      await output.copyPages(
        source,
        [i]
      );

    output.addPage(page);

    result.push({
      name: `page-${i + 1}.pdf`,
      blob: new Blob(
        [await output.save()],
        {
          type: "application/pdf",
        }
      ),
    });
  }

  return result;
}

/* =========================================================
   COMPRESS PDF
========================================================= */

async function compressPDF(file) {

  const pdf =
    await PDFDocument.load(
      await file.arrayBuffer()
    );

  const bytes =
    await pdf.save({
      useObjectStreams: true,
      addDefaultPage: false,
    });

  return new Blob(
    [bytes],
    {
      type: "application/pdf",
    }
  );
}

/* =========================================================
   JPG TO PDF
========================================================= */

async function imageToPDF(files) {

  const pdf =
    new jsPDF({
      unit: "pt",
      format: "a4",
    });

  for (
    let i = 0;
    i < files.length;
    i++
  ) {

    if (i > 0) {
      pdf.addPage();
    }

    const data =
      await fileData(files[i]);

    const img =
      new window.Image();

    await new Promise(
      (resolve, reject) => {

        img.onload = resolve;
        img.onerror = reject;

        img.src = data;
      }
    );

    const pageWidth = 595;
    const pageHeight = 842;
    const margin = 35;

    const maxWidth =
      pageWidth - margin * 2;

    const maxHeight =
      pageHeight - margin * 2;

    const ratio =
      Math.min(
        maxWidth / img.width,
        maxHeight / img.height
      );

    const imageType =
      files[i].type === "image/png"
        ? "PNG"
        : "JPEG";

    pdf.addImage(
      data,
      imageType,
      margin,
      margin,
      img.width * ratio,
      img.height * ratio
    );
  }

  return pdf.output("blob");
}

/* =========================================================
   PDF TO JPG
========================================================= */

async function pdfToJPG(file) {

  const pdf =
    await pdfjsLib
      .getDocument({
        data:
          await file.arrayBuffer(),
      })
      .promise;

  const result = [];

  for (
    let pageNo = 1;
    pageNo <= pdf.numPages;
    pageNo++
  ) {

    const page =
      await pdf.getPage(pageNo);

    const viewport =
      page.getViewport({
        scale: 2,
      });

    const canvas =
      document.createElement(
        "canvas"
      );

    canvas.width =
      viewport.width;

    canvas.height =
      viewport.height;

    const context =
      canvas.getContext("2d");

    await page.render({
      canvasContext: context,
      viewport,
    }).promise;

    const blob =
      await new Promise(
        (resolve) => {

          canvas.toBlob(
            resolve,
            "image/jpeg",
            0.92
          );

        }
      );

    result.push({
      name:
        `page-${pageNo}.jpg`,
      blob,
    });
  }

  return result;
}

/* =========================================================
   PDF TO WORD
========================================================= */

async function pdfToWord(file) {

  const pdf =
    await pdfjsLib
      .getDocument({
        data: await file.arrayBuffer(),
      })
      .promise;

  const paragraphs = [];

  for (
    let pageNo = 1;
    pageNo <= pdf.numPages;
    pageNo++
  ) {

    const page =
      await pdf.getPage(pageNo);

    const textContent =
      await page.getTextContent();

    const items =
      textContent.items
        .filter(
          (item) =>
            typeof item.str === "string" &&
            item.str.trim()
        )
        .map((item) => ({
          text: item.str.trim(),
          x: item.transform?.[4] ?? 0,
          y: item.transform?.[5] ?? 0,
        }))
        .sort((a, b) => {

          const yDiff =
            Math.abs(
              b.y - a.y
            );

          return yDiff > 3
            ? b.y - a.y
            : a.x - b.x;
        });

    const lines = [];

    for (const item of items) {

      let line =
        lines.find(
          (candidate) =>
            Math.abs(
              candidate.y - item.y
            ) <= 3
        );

      if (!line) {

        line = {
          y: item.y,
          items: [],
        };

        lines.push(line);
      }

      line.items.push(item);
    }

    lines.sort(
      (a, b) =>
        b.y - a.y
    );

    if (pageNo > 1) {

      paragraphs.push(
        new Paragraph({
          children: [
            new TextRun({
              text: "",
            }),
          ],
          pageBreakBefore: true,
        })
      );
    }

    for (const line of lines) {

      line.items.sort(
        (a, b) =>
          a.x - b.x
      );

      const text =
        line.items
          .map(
            (item) =>
              item.text
          )
          .join(" ")
          .replace(/\s+/g, " ")
          .trim();

      if (text) {

        paragraphs.push(
          new Paragraph({
            children: [
              new TextRun({
                text,
              }),
            ],
            spacing: {
              after: 120,
            },
          })
        );
      }
    }
  }

  if (!paragraphs.length) {

    throw new Error(
      "No selectable text was found. This may be a scanned PDF."
    );
  }

  const document =
    new Document({
      sections: [
        {
          properties: {},
          children: paragraphs,
        },
      ],
    });

  return Packer.toBlob(
    document
  );
}

/* =========================================================
   WORD TO PDF
========================================================= */

async function wordToPDF(file) {

  const zip =
    await JSZip.loadAsync(
      await file.arrayBuffer()
    );

  const documentXml =
    zip.file(
      "word/document.xml"
    );

  if (!documentXml) {

    throw new Error(
      "Invalid Word document."
    );
  }

  const xml =
    await documentXml.async(
      "text"
    );

  const parser =
    new DOMParser();

  const doc =
    parser.parseFromString(
      xml,
      "application/xml"
    );

  const paragraphs =
    [...doc.getElementsByTagName("w:p")];

  const pdf =
    new jsPDF({
      unit: "pt",
      format: "a4",
    });

  const margin = 45;
  const pageWidth = 595;
  const pageHeight = 842;

  let y = 55;

  for (const paragraph of paragraphs) {

    const textNodes =
      [...paragraph.getElementsByTagName("w:t")];

    const text =
      textNodes
        .map(
          (node) =>
            node.textContent || ""
        )
        .join("")
        .trim();

    if (!text) {
      y += 12;
      continue;
    }

    const lines =
      pdf.splitTextToSize(
        text,
        pageWidth - margin * 2
      );

    for (const line of lines) {

      if (y > pageHeight - 55) {

        pdf.addPage();
        y = 55;
      }

      pdf.setFont(
        "helvetica",
        "normal"
      );

      pdf.setFontSize(11);

      pdf.text(
        line,
        margin,
        y
      );

      y += 16;
    }

    y += 8;
  }

  return pdf.output(
    "blob"
  );
}

/* =========================================================
   PDF TO POWERPOINT
========================================================= */

async function pdfToPowerPoint(file) {

  const pdf =
    await pdfjsLib
      .getDocument({
        data: await file.arrayBuffer(),
      })
      .promise;

  const pptx =
    new PptxGenJS();

  pptx.layout =
    "LAYOUT_WIDE";

  pptx.author =
    "DesiPDF";

  pptx.subject =
    "PDF converted to PowerPoint";

  pptx.title =
    "DesiPDF Presentation";

  for (
    let pageNo = 1;
    pageNo <= pdf.numPages;
    pageNo++
  ) {

    const page =
      await pdf.getPage(pageNo);

    const viewport =
      page.getViewport({
        scale: 2,
      });

    const canvas =
      document.createElement(
        "canvas"
      );

    canvas.width =
      Math.ceil(viewport.width);

    canvas.height =
      Math.ceil(viewport.height);

    const context =
      canvas.getContext("2d");

    await page.render({
      canvasContext: context,
      viewport,
    }).promise;

    const imageData =
      canvas.toDataURL(
        "image/jpeg",
        0.92
      );

    const slide =
      pptx.addSlide();

    slide.background = {
      color: "FFFFFF",
    };

    const slideWidth =
      13.333;

    const slideHeight =
      7.5;

    const pdfWidth =
      viewport.width;

    const pdfHeight =
      viewport.height;

    const scale =
      Math.min(
        slideWidth / pdfWidth,
        slideHeight / pdfHeight
      );

    const imageWidth =
      pdfWidth * scale;

    const imageHeight =
      pdfHeight * scale;

    const x =
      (slideWidth - imageWidth) / 2;

    const y =
      (slideHeight - imageHeight) / 2;

    slide.addImage({
      data: imageData,
      x,
      y,
      w: imageWidth,
      h: imageHeight,
    });
  }

  return pptx.write({
    outputType: "blob",
  });
}

/* =========================================================
   POWERPOINT TO PDF
========================================================= */

async function powerPointToPDF(file) {

  const zip =
    await JSZip.loadAsync(
      await file.arrayBuffer()
    );

  const slideFiles =
    Object.keys(zip.files)
      .filter(
        (name) =>
          /^ppt\/slides\/slide\d+\.xml$/.test(
            name
          )
      )
      .sort((a, b) => {

        const na =
          Number(
            a.match(/slide(\d+)\.xml/)?.[1] || 0
          );

        const nb =
          Number(
            b.match(/slide(\d+)\.xml/)?.[1] || 0
          );

        return na - nb;
      });

  if (!slideFiles.length) {

    throw new Error(
      "No PowerPoint slides were found."
    );
  }

  const pdf =
    new jsPDF({
      unit: "pt",
      format: "a4",
      orientation: "landscape",
    });

  for (
    let i = 0;
    i < slideFiles.length;
    i++
  ) {

    if (i > 0) {
      pdf.addPage();
    }

    const xml =
      await zip.files[
        slideFiles[i]
      ].async("text");

    const parser =
      new DOMParser();

    const doc =
      parser.parseFromString(
        xml,
        "application/xml"
      );

    const textNodes =
      [...doc.getElementsByTagName("a:t")];

    const text =
      textNodes
        .map(
          (node) =>
            node.textContent || ""
        )
        .join(" ")
        .replace(/\s+/g, " ")
        .trim();

    pdf.setFont(
      "helvetica",
      "normal"
    );

    pdf.setFontSize(18);

    pdf.text(
      `Slide ${i + 1}`,
      40,
      45
    );

    pdf.setFontSize(12);

    const lines =
      pdf.splitTextToSize(
        text || "PowerPoint slide",
        690
      );

    pdf.text(
      lines,
      40,
      80
    );
  }

  return pdf.output(
    "blob"
  );
}

/* =========================================================
   EXCEL TO PDF
========================================================= */

async function excelToPDF(file) {

  const workbook =
    XLSX.read(
      await file.arrayBuffer(),
      {
        type: "array",
      }
    );

  const pdf =
    new jsPDF({
      unit: "pt",
      format: "a4",
      orientation: "landscape",
    });

  const pageWidth =
    842;

  const pageHeight =
    595;

  const margin =
    30;

  workbook.SheetNames.forEach(
    (sheetName, sheetIndex) => {

      if (sheetIndex > 0) {
        pdf.addPage();
      }

      const sheet =
        workbook.Sheets[
          sheetName
        ];

      const rows =
        XLSX.utils.sheet_to_json(
          sheet,
          {
            header: 1,
            raw: false,
          }
        );

      pdf.setFontSize(16);

      pdf.text(
        sheetName,
        margin,
        30
      );

      let y = 55;

      const maxCols =
        Math.min(
          10,
          Math.max(
            1,
            ...rows.map(
              (row) =>
                row.length
            )
          )
        );

      const tableWidth =
        pageWidth - margin * 2;

      const colWidth =
        tableWidth / maxCols;

      for (
        let r = 0;
        r < rows.length;
        r++
      ) {

        if (
          y >
          pageHeight - 30
        ) {

          pdf.addPage();
          y = 35;
        }

        const row =
          rows[r] || [];

        for (
          let c = 0;
          c < maxCols;
          c++
        ) {

          const value =
            row[c] == null
              ? ""
              : String(row[c]);

          const x =
            margin +
            c * colWidth;

          pdf.setFontSize(
            r === 0 ? 8 : 7
          );

          pdf.rect(
            x,
            y,
            colWidth,
            18
          );

          const shortText =
            value.length > 28
              ? value.slice(0, 25) + "..."
              : value;

          pdf.text(
            shortText,
            x + 3,
            y + 12
          );
        }

        y += 18;
      }
    }
  );

  return pdf.output(
    "blob"
  );
}

/* =========================================================
   PDF TO EXCEL
========================================================= */

async function pdfToExcel(file) {

  const pdf =
    await pdfjsLib
      .getDocument({
        data:
          await file.arrayBuffer(),
      })
      .promise;

  const rows = [];

  for (
    let pageNo = 1;
    pageNo <= pdf.numPages;
    pageNo++
  ) {

    const page =
      await pdf.getPage(
        pageNo
      );

    const content =
      await page.getTextContent();

    const items =
      content.items
        .filter(
          (item) =>
            typeof item.str === "string" &&
            item.str.trim()
        )
        .map((item) => ({
          text:
            item.str.trim(),
          x:
            item.transform?.[4] ?? 0,
          y:
            item.transform?.[5] ?? 0,
        }));

    items.sort((a, b) => {

      const yDiff =
        Math.abs(
          b.y - a.y
        );

      if (yDiff > 4) {
        return b.y - a.y;
      }

      return a.x - b.x;
    });

    let currentY = null;
    let currentRow = [];

    for (const item of items) {

      if (
        currentY === null ||
        Math.abs(
          currentY - item.y
        ) <= 4
      ) {

        currentRow.push(
          item.text
        );

        if (
          currentY === null
        ) {
          currentY = item.y;
        }

      } else {

        rows.push(
          currentRow
        );

        currentRow = [
          item.text,
        ];

        currentY = item.y;
      }
    }

    if (currentRow.length) {
      rows.push(
        currentRow
      );
    }

    rows.push([]);
  }

  const worksheet =
    XLSX.utils.aoa_to_sheet(
      rows
    );

  const workbook =
    XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(
    workbook,
    worksheet,
    "PDF Data"
  );

  const arrayBuffer =
    XLSX.write(
      workbook,
      {
        bookType: "xlsx",
        type: "array",
      }
    );

  return new Blob(
    [arrayBuffer],
    {
      type:
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    }
  );
}

/* =========================================================
   HTML TO PDF
========================================================= */

async function htmlToPDF(file) {

  const html =
    await file.text();

  const container =
    document.createElement(
      "div"
    );

  container.style.position =
    "fixed";

  container.style.left =
    "-100000px";

  container.style.top =
    "0";

  container.style.width =
    "794px";

  container.style.background =
    "#ffffff";

  container.style.padding =
    "30px";

  container.innerHTML =
    html;

  document.body.appendChild(
    container
  );

  try {

    const canvas =
      await html2canvas(
        container,
        {
          scale: 1.5,
          useCORS: true,
          backgroundColor:
            "#ffffff",
        }
      );

    const imgData =
      canvas.toDataURL(
        "image/jpeg",
        0.92
      );

    const pdf =
      new jsPDF({
        unit: "pt",
        format: "a4",
      });

    const pageWidth =
      595;

    const pageHeight =
      842;

    const ratio =
      pageWidth /
      canvas.width;

    const imgWidth =
      pageWidth;

    const imgHeight =
      canvas.height * ratio;

    let heightLeft =
      imgHeight;

    let position =
      0;

    pdf.addImage(
      imgData,
      "JPEG",
      0,
      position,
      imgWidth,
      imgHeight
    );

    heightLeft -=
      pageHeight;

    while (
      heightLeft > 0
    ) {

      position -=
        pageHeight;

      pdf.addPage();

      pdf.addImage(
        imgData,
        "JPEG",
        0,
        position,
        imgWidth,
        imgHeight
      );

      heightLeft -=
        pageHeight;
    }

    return pdf.output(
      "blob"
    );

  } finally {

    container.remove();
  }
}

/* =========================================================
   PDF TO PDF/A
========================================================= */

async function pdfToPDFA(file) {

  /*
    Browser-only pdf-lib does not guarantee full PDF/A
    standards compliance. This creates a fresh archival-style
    PDF copy while preserving page appearance.
  */

  const pdf =
    await pdfjsLib
      .getDocument({
        data:
          await file.arrayBuffer(),
      })
      .promise;

  const output =
    new jsPDF({
      unit: "pt",
      format: "a4",
    });

  for (
    let pageNo = 1;
    pageNo <= pdf.numPages;
    pageNo++
  ) {

    if (pageNo > 1) {
      output.addPage();
    }

    const page =
      await pdf.getPage(
        pageNo
      );

    const viewport =
      page.getViewport({
        scale: 1.5,
      });

    const canvas =
      document.createElement(
        "canvas"
      );

    canvas.width =
      Math.ceil(
        viewport.width
      );

    canvas.height =
      Math.ceil(
        viewport.height
      );

    const context =
      canvas.getContext(
        "2d"
      );

    await page.render({
      canvasContext:
        context,
      viewport,
    }).promise;

    const image =
      canvas.toDataURL(
        "image/jpeg",
        0.95
      );

    const pageWidth =
      595;

    const pageHeight =
      842;

    const ratio =
      Math.min(
        pageWidth /
          viewport.width,
        pageHeight /
          viewport.height
      );

    const w =
      viewport.width *
      ratio;

    const h =
      viewport.height *
      ratio;

    const x =
      (pageWidth - w) / 2;

    const y =
      (pageHeight - h) / 2;

    output.addImage(
      image,
      "JPEG",
      x,
      y,
      w,
      h
    );
  }

  output.setProperties({
    title:
      "DesiPDF PDF/A Archive",
    subject:
      "Archived PDF copy",
    author:
      "DesiPDF",
    creator:
      "DesiPDF",
    keywords:
      "PDF/A, archive, DesiPDF",
  });

  return output.output(
    "blob"
  );
}

/* =========================================================
   ROTATE PDF
========================================================= */

async function rotatePDF(file) {

  const pdf =
    await PDFDocument.load(
      await file.arrayBuffer()
    );

  pdf.getPages().forEach(
    (page) => {

      const current =
        page.getRotation().angle;

      page.setRotation(
        degrees(
          (current + 90) % 360
        )
      );
    }
  );

  return new Blob(
    [await pdf.save()],
    {
      type: "application/pdf",
    }
  );
}

/* =========================================================
   REMOVE PAGES
========================================================= */

async function removePages(
  file,
  pagesToRemove
) {

  const pdf =
    await PDFDocument.load(
      await file.arrayBuffer()
    );

  const total =
    pdf.getPageCount();

  const pages =
    [...new Set(pagesToRemove)]
      .filter(
        (page) =>
          page >= 0 &&
          page < total
      )
      .sort(
        (a, b) =>
          b - a
      );

  if (
    pages.length >= total
  ) {

    throw new Error(
      "You cannot remove every page."
    );
  }

  pages.forEach(
    (page) =>
      pdf.removePage(
        page
      )
  );

  return new Blob(
    [await pdf.save()],
    {
      type: "application/pdf",
    }
  );
}

/* =========================================================
   EXTRACT PAGES
========================================================= */

async function extractPages(
  file,
  pageNumbers
) {

  const source =
    await PDFDocument.load(
      await file.arrayBuffer()
    );

  const total =
    source.getPageCount();

  const indexes =
    [...new Set(pageNumbers)]
      .filter(
        (page) =>
          page >= 0 &&
          page < total
      );

  if (!indexes.length) {

    throw new Error(
      "Please enter valid page numbers."
    );
  }

  const output =
    await PDFDocument.create();

  const pages =
    await output.copyPages(
      source,
      indexes
    );

  pages.forEach(
    (page) =>
      output.addPage(page)
  );

  return new Blob(
    [await output.save()],
    {
      type: "application/pdf",
    }
  );
}

/* =========================================================
   REORDER PAGES
========================================================= */

async function reorderPages(
  file,
  order
) {

  const source =
    await PDFDocument.load(
      await file.arrayBuffer()
    );

  const total =
    source.getPageCount();

  if (
    order.length !== total
  ) {

    throw new Error(
      `Please enter all ${total} page numbers in the new order.`
    );
  }

  const indexes =
    order.map(
      (page) =>
        page - 1
    );

  const unique =
    new Set(indexes);

  if (
    unique.size !== total ||
    indexes.some(
      (i) =>
        i < 0 ||
        i >= total
    )
  ) {

    throw new Error(
      "Every page number must appear exactly once."
    );
  }

  const output =
    await PDFDocument.create();

  const pages =
    await output.copyPages(
      source,
      indexes
    );

  pages.forEach(
    (page) =>
      output.addPage(page)
  );

  return new Blob(
    [await output.save()],
    {
      type: "application/pdf",
    }
  );
}

/* =========================================================
   WATERMARK
========================================================= */

async function watermarkPDF(
  file,
  text
) {

  const pdf =
    await PDFDocument.load(
      await file.arrayBuffer()
    );

  for (
    const page of pdf.getPages()
  ) {

    const {
      width,
      height,
    } = page.getSize();

    page.drawText(
      text || "DesiPDF",
      {
        x:
          width / 2 - 100,

        y:
          height / 2,

        size: 34,

        color:
          rgb(
            0.5,
            0.5,
            0.5
          ),

        opacity: 0.25,

        rotate:
          degrees(45),
      }
    );
  }

  return new Blob(
    [await pdf.save()],
    {
      type: "application/pdf",
    }
  );
}

/* =========================================================
   PAGE NUMBERS
========================================================= */

async function addPageNumbers(
  file
) {

  const pdf =
    await PDFDocument.load(
      await file.arrayBuffer()
    );

  const font =
    await pdf.embedFont(
      StandardFonts.Helvetica
    );

  pdf.getPages().forEach(
    (page, index) => {

      const {
        width,
      } = page.getSize();

      const text =
        String(
          index + 1
        );

      page.drawText(
        text,
        {
          x:
            width / 2 - 5,

          y: 20,

          size: 12,

          font,

          color:
            rgb(
              0.2,
              0.2,
              0.2
            ),
        }
      );
    }
  );

  return new Blob(
    [await pdf.save()],
    {
      type: "application/pdf",
    }
  );
}

/* =========================================================
   PROTECT PDF
========================================================= */

async function protectPDF(
  file,
  password,
  ownerPassword,
  options
) {

  if (!password) {

    throw new Error(
      "Please enter a password."
    );
  }

  const bytes =
    new Uint8Array(
      await file.arrayBuffer()
    );

  const encrypted =
    await encryptPDF(
      bytes,
      password,
      {
        ownerPassword:
          ownerPassword ||
          password,

        algorithm:
          "AES-256",

        allowPrinting:
          options.allowPrinting,

        allowModifying:
          options.allowModifying,

        allowCopying:
          options.allowCopying,

        allowAnnotating:
          options.allowAnnotating,

        allowFillingForms:
          true,

        allowExtraction:
          options.allowExtraction,

        allowAssembly:
          options.allowAssembly,

        allowHighQualityPrint:
          options.allowHighQualityPrint,
      }
    );

  return new Blob(
    [encrypted],
    {
      type: "application/pdf",
    }
  );
}

/* =========================================================
   UNLOCK PDF
========================================================= */

async function unlockPDF(
  file,
  password
) {

  if (!password) {

    throw new Error(
      "Please enter the PDF password."
    );
  }

  const bytes =
    new Uint8Array(
      await file.arrayBuffer()
    );

  const info =
    await isEncrypted(
      bytes
    );

  if (!info.encrypted) {

    throw new Error(
      "This PDF is not encrypted."
    );
  }

  const decrypted =
    await decryptPDF(
      bytes,
      password
    );

  return new Blob(
    [decrypted],
    {
      type: "application/pdf",
    }
  );
}

/* =========================================================
   SIGN PDF
========================================================= */

async function signPDF(
  file,
  signature,
  pageNumber,
  xPosition,
  yPosition
) {

  if (!signature.trim()) {

    throw new Error(
      "Please enter your signature."
    );
  }

  const pdf =
    await PDFDocument.load(
      await file.arrayBuffer()
    );

  const pages =
    pdf.getPages();

  const pageIndex =
    Math.max(
      0,
      Math.min(
        pages.length - 1,
        Number(
          pageNumber || 1
        ) - 1
      )
    );

  const page =
    pages[pageIndex];

  const font =
    await pdf.embedFont(
      StandardFonts.HelveticaOblique
    );

  const {
    width,
    height,
  } = page.getSize();

  const x =
    Number.isFinite(
      Number(xPosition)
    )
      ? Number(xPosition)
      : 60;

  const y =
    Number.isFinite(
      Number(yPosition)
    )
      ? Number(yPosition)
      : height - 100;

  page.drawText(
    signature,
    {
      x:
        Math.max(
          10,
          Math.min(
            width - 200,
            x
          )
        ),

      y:
        Math.max(
          10,
          Math.min(
            height - 40,
            y
          )
        ),

      size: 28,

      font,

      color:
        rgb(
          0.05,
          0.1,
          0.5
        ),
    }
  );

  return new Blob(
    [await pdf.save()],
    {
      type: "application/pdf",
    }
  );
}

/* =========================================================
   PAGE NUMBER PARSER
========================================================= */

function parsePageNumbers(
  value
) {

  if (!value.trim()) {

    throw new Error(
      "Please enter page numbers."
    );
  }

  const pages =
    value
      .split(",")
      .map(
        (item) =>
          Number(
            item.trim()
          )
      );

  if (
    pages.some(
      (page) =>
        !Number.isInteger(page) ||
        page < 1
    )
  ) {

    throw new Error(
      "Page numbers must be positive numbers."
    );
  }

  return pages;
}

/* =========================================================
   TOOL PAGE
========================================================= */

function ToolPage() {

  const { id } =
    useParams();
    useEffect(() => {
  const seo = seoData[id];

  if (!seo) return;

  document.title = seo.title;

  let descriptionTag = document.querySelector(
    'meta[name="description"]'
  );

  if (!descriptionTag) {
    descriptionTag = document.createElement("meta");
    descriptionTag.setAttribute("name", "description");
    document.head.appendChild(descriptionTag);
  }

  descriptionTag.setAttribute(
    "content",
    seo.description
  );

  let canonicalTag = document.querySelector(
    'link[rel="canonical"]'
  );

  if (!canonicalTag) {
    canonicalTag = document.createElement("link");
    canonicalTag.setAttribute("rel", "canonical");
    document.head.appendChild(canonicalTag);
  }

  canonicalTag.setAttribute(
    "href",
    `https://desipdf.online/tool/${id}`
  );
}, [id]);

  const tool =
    tools.find(
      (item) =>
        item.id === id
    ) || tools[0];

  const [files, setFiles] =
    useState([]);

  const [busy, setBusy] =
    useState(false);

  const [done, setDone] =
    useState(false);

  const [msg, setMsg] =
    useState("");

  const [pageInput, setPageInput] =
    useState("");

  const [watermark, setWatermark] =
    useState("DesiPDF");

  const [password, setPassword] =
    useState("");

  const [ownerPassword, setOwnerPassword] =
    useState("");

  const [signature, setSignature] =
    useState("");

  const [signaturePage, setSignaturePage] =
    useState("1");

  const [signatureX, setSignatureX] =
    useState("60");

  const [signatureY, setSignatureY] =
    useState("700");

  const [permissions, setPermissions] =
    useState({
      allowPrinting: true,
      allowModifying: false,
      allowCopying: false,
      allowAnnotating: true,
      allowExtraction: false,
      allowAssembly: false,
      allowHighQualityPrint: true,
    });

  /* =======================================================
     MULTIPLE FILE SETTINGS
  ======================================================= */

  const isMultiple =
    id === "merge-pdf" ||
    id === "jpg-to-pdf";

  /* =======================================================
     ACCEPT FILE TYPES
  ======================================================= */

  let accept = ".pdf";

  if (
    id === "jpg-to-pdf"
  ) {
    accept =
      "image/jpeg,image/png,.jpg,.jpeg,.png";
  }

  if (
    id === "word-to-pdf"
  ) {
    accept =
      ".docx";
  }

  if (
    id === "powerpoint-to-pdf"
  ) {
    accept =
      ".pptx";
  }

  if (
    id === "excel-to-pdf"
  ) {
    accept =
      ".xlsx,.xls,.csv";
  }

  if (
    id === "html-to-pdf"
  ) {
    accept =
      ".html,.htm";
  }

  /* =======================================================
     RUN
  ======================================================= */

  async function run() {

    if (!files.length) {

      setMsg(
        "Please select a file first."
      );

      return;
    }

    if (
      id === "merge-pdf" &&
      files.length < 2
    ) {

      setMsg(
        "Merge PDF needs at least two PDF files."
      );

      return;
    }

    setBusy(true);
    setDone(false);
    setMsg("");

    try {

      let blob;

      /* MERGE */

      if (
        id === "merge-pdf"
      ) {

        blob =
          await mergePDF(
            files
          );

        download(
          blob,
          "merged.pdf"
        );
      }

      /* SPLIT */

      else if (
        id === "split-pdf"
      ) {

        const pages =
          await splitPDF(
            files[0]
          );

        for (
          const page of pages
        ) {

          download(
            page.blob,
            page.name
          );

          await sleep(
            250
          );
        }
      }

      /* COMPRESS */

      else if (
        id === "compress-pdf"
      ) {

        blob =
          await compressPDF(
            files[0]
          );

        download(
          blob,
          "compressed.pdf"
        );
      }

      /* JPG TO PDF */

      else if (
        id === "jpg-to-pdf"
      ) {

        blob =
          await imageToPDF(
            files
          );

        download(
          blob,
          "images.pdf"
        );
      }

      /* PDF TO JPG */

      else if (
        id === "pdf-to-jpg"
      ) {

        const images =
          await pdfToJPG(
            files[0]
          );

        for (
          const image of images
        ) {

          download(
            image.blob,
            image.name
          );

          await sleep(
            250
          );
        }
      }

      /* WORD TO PDF */

      else if (
        id === "word-to-pdf"
      ) {

        blob =
          await wordToPDF(
            files[0]
          );

        download(
          blob,
          "converted.pdf"
        );
      }

      /* PDF TO WORD */

      else if (
        id === "pdf-to-word"
      ) {

        blob =
          await pdfToWord(
            files[0]
          );

        download(
          blob,
          "converted.docx"
        );
      }

      /* POWERPOINT TO PDF */

      else if (
        id === "powerpoint-to-pdf"
      ) {

        blob =
          await powerPointToPDF(
            files[0]
          );

        download(
          blob,
          "converted.pdf"
        );
      }

      /* PDF TO POWERPOINT */

      else if (
        id === "pdf-to-powerpoint"
      ) {

        blob =
          await pdfToPowerPoint(
            files[0]
          );

        download(
          blob,
          "converted.pptx"
        );
      }

      /* EXCEL TO PDF */

      else if (
        id === "excel-to-pdf"
      ) {

        blob =
          await excelToPDF(
            files[0]
          );

        download(
          blob,
          "converted.pdf"
        );
      }

      /* PDF TO EXCEL */

      else if (
        id === "pdf-to-excel"
      ) {

        blob =
          await pdfToExcel(
            files[0]
          );

        download(
          blob,
          "converted.xlsx"
        );
      }

      /* HTML TO PDF */

      else if (
        id === "html-to-pdf"
      ) {

        blob =
          await htmlToPDF(
            files[0]
          );

        download(
          blob,
          "converted.pdf"
        );
      }

      /* PDF TO PDF/A */

      else if (
        id === "pdf-to-pdfa"
      ) {

        blob =
          await pdfToPDFA(
            files[0]
          );

        download(
          blob,
          "archived-pdfa.pdf"
        );
      }

      /* ROTATE */

      else if (
        id === "rotate-pdf"
      ) {

        blob =
          await rotatePDF(
            files[0]
          );

        download(
          blob,
          "rotated.pdf"
        );
      }

      /* REMOVE */

      else if (
        id === "remove-pages"
      ) {

        const pages =
          parsePageNumbers(
            pageInput
          );

        blob =
          await removePages(
            files[0],
            pages.map(
              (x) =>
                x - 1
            )
          );

        download(
          blob,
          "pages-removed.pdf"
        );
      }

      /* EXTRACT */

      else if (
        id === "extract-pages"
      ) {

        const pages =
          parsePageNumbers(
            pageInput
          );

        blob =
          await extractPages(
            files[0],
            pages.map(
              (x) =>
                x - 1
            )
          );

        download(
          blob,
          "extracted-pages.pdf"
        );
      }

      /* REORDER */

      else if (
        id === "reorder-pages"
      ) {

        const pages =
          parsePageNumbers(
            pageInput
          );

        blob =
          await reorderPages(
            files[0],
            pages
          );

        download(
          blob,
          "reordered.pdf"
        );
      }

      /* WATERMARK */

      else if (
        id === "watermark-pdf"
      ) {

        blob =
          await watermarkPDF(
            files[0],
            watermark
          );

        download(
          blob,
          "watermarked.pdf"
        );
      }

      /* PAGE NUMBERS */

      else if (
        id === "page-numbers"
      ) {

        blob =
          await addPageNumbers(
            files[0]
          );

        download(
          blob,
          "numbered.pdf"
        );
      }

      /* PROTECT */

      else if (
        id === "protect-pdf"
      ) {

        blob =
          await protectPDF(
            files[0],
            password,
            ownerPassword,
            permissions
          );

        download(
          blob,
          "protected.pdf"
        );
      }

      /* UNLOCK */

      else if (
        id === "unlock-pdf"
      ) {

        blob =
          await unlockPDF(
            files[0],
            password
          );

        download(
          blob,
          "unlocked.pdf"
        );
      }

      /* SIGN */

      else if (
        id === "sign-pdf"
      ) {

        blob =
          await signPDF(
            files[0],
            signature,
            signaturePage,
            signatureX,
            signatureY
          );

        download(
          blob,
          "signed.pdf"
        );
      }

      setDone(true);

    } catch (error) {

      console.error(
        error
      );

      setMsg(
        error?.message ||
        "Something went wrong while processing the file."
      );

    } finally {

      setBusy(false);
    }
  }

  return (
    <>
      <Header />

      <main className="toolPage">

        <div className="container">

          <Link
            to="/"
            className="back"
          >
            ← All tools
          </Link>

          <div className="toolHero">

            <div className="bigIcon">
              {React.createElement(
                tool.icon,
                {
                  size: 30,
                }
              )}
            </div>

            <span className="eyebrow">
              {tool.cat}
            </span>

            <h1>
              {tool.name}
            </h1>

            <p>
              {tool.desc}
            </p>

          </div>

          <div className="workspace">

            <FileDrop
              multiple={isMultiple}
              accept={accept}
              onFiles={setFiles}
            />

            {/* FILE LIST */}

            {files.length > 0 && (

              <div className="fileList">

                {files.map(
                  (file, index) => (

                    <div
                      className="file"
                      key={`${file.name}-${index}`}
                    >

                      <FileText size={19} />

                      <span>
                        {file.name}
                      </span>

                      <small>
                        {(
                          file.size /
                          1024 /
                          1024
                        ).toFixed(2)}
                        {" "}MB
                      </small>

                    </div>

                  )
                )}

              </div>

            )}

            {/* PAGE INPUT */}

            {[
              "remove-pages",
              "extract-pages",
              "reorder-pages",
            ].includes(id) &&
              files.length > 0 && (

                <div className="removeBox">

                  <p>
                    {id === "reorder-pages"
                      ? "Enter the complete new page order. Example: 3,1,2,4"
                      : "Enter page numbers separated by commas. Example: 2,4,6"}
                  </p>

                  <input
                    value={pageInput}
                    onChange={(e) =>
                      setPageInput(
                        e.target.value
                      )
                    }
                    placeholder={
                      id === "reorder-pages"
                        ? "3,1,2,4"
                        : "2,4,6"
                    }
                  />

                </div>

              )}

            {/* WATERMARK */}

            {id ===
              "watermark-pdf" &&
              files.length > 0 && (

                <div className="removeBox">

                  <p>
                    Watermark text
                  </p>

                  <input
                    value={watermark}
                    onChange={(e) =>
                      setWatermark(
                        e.target.value
                      )
                    }
                    placeholder="DesiPDF"
                  />

                </div>

              )}

            {/* PASSWORD */}

            {[
              "protect-pdf",
              "unlock-pdf",
            ].includes(id) &&
              files.length > 0 && (

                <div className="securityBox">

                  <h3>
                    {id === "protect-pdf"
                      ? "PDF Security"
                      : "PDF Password"}
                  </h3>

                  <input
                    type="password"
                    value={password}
                    onChange={(e) =>
                      setPassword(
                        e.target.value
                      )
                    }
                    placeholder={
                      id === "protect-pdf"
                        ? "Enter password"
                        : "Enter PDF password"
                    }
                  />

                  {id === "protect-pdf" && (

                    <>

                      <input
                        type="password"
                        value={ownerPassword}
                        onChange={(e) =>
                          setOwnerPassword(
                            e.target.value
                          )
                        }
                        placeholder="Owner password (optional)"
                      />

                      <div className="permissions">

                        <label>
                          <input
                            type="checkbox"
                            checked={
                              permissions.allowPrinting
                            }
                            onChange={(e) =>
                              setPermissions({
                                ...permissions,
                                allowPrinting:
                                  e.target.checked,
                              })
                            }
                          />
                          Allow printing
                        </label>

                        <label>
                          <input
                            type="checkbox"
                            checked={
                              permissions.allowCopying
                            }
                            onChange={(e) =>
                              setPermissions({
                                ...permissions,
                                allowCopying:
                                  e.target.checked,
                              })
                            }
                          />
                          Allow copying
                        </label>

                        <label>
                          <input
                            type="checkbox"
                            checked={
                              permissions.allowModifying
                            }
                            onChange={(e) =>
                              setPermissions({
                                ...permissions,
                                allowModifying:
                                  e.target.checked,
                              })
                            }
                          />
                          Allow modifying
                        </label>

                        <label>
                          <input
                            type="checkbox"
                            checked={
                              permissions.allowAnnotating
                            }
                            onChange={(e) =>
                              setPermissions({
                                ...permissions,
                                allowAnnotating:
                                  e.target.checked,
                              })
                            }
                          />
                          Allow annotations
                        </label>

                      </div>

                    </>
                  )}

                </div>

              )}

            {/* SIGN */}

            {id ===
              "sign-pdf" &&
              files.length > 0 && (

                <div className="securityBox">

                  <h3>
                    Add Signature
                  </h3>

                  <input
                    value={signature}
                    onChange={(e) =>
                      setSignature(
                        e.target.value
                      )
                    }
                    placeholder="Type your signature"
                  />

                  <input
                    type="number"
                    min="1"
                    value={signaturePage}
                    onChange={(e) =>
                      setSignaturePage(
                        e.target.value
                      )
                    }
                    placeholder="Page number"
                  />

                  <div
                    style={{
                      display: "grid",
                      gridTemplateColumns:
                        "1fr 1fr",
                      gap: 12,
                    }}
                  >

                    <input
                      type="number"
                      value={signatureX}
                      onChange={(e) =>
                        setSignatureX(
                          e.target.value
                        )
                      }
                      placeholder="X position"
                    />

                    <input
                      type="number"
                      value={signatureY}
                      onChange={(e) =>
                        setSignatureY(
                          e.target.value
                        )
                      }
                      placeholder="Y position"
                    />

                  </div>

                  <p>
                    Signature position uses PDF coordinates.
                  </p>

                </div>

              )}

            {/* CONVERSION INFO */}

            {[
              "word-to-pdf",
              "powerpoint-to-pdf",
              "excel-to-pdf",
              "html-to-pdf",
              "pdf-to-pdfa",
            ].includes(id) &&
              files.length > 0 && (

                <div className="securityBox">

                  <h3>
                    {tool.name}
                  </h3>

                  <p>
                    Your file will be processed
                    directly in the browser.
                  </p>

                  {id === "pdf-to-pdfa" && (
                    <p>
                      The browser creates an
                      archival-style PDF copy.
                    </p>
                  )}

                </div>

              )}

            {/* ERROR */}

            {msg && (

              <div className="notice">
                {msg}
              </div>

            )}

            {/* PROCESS */}

            <button
              className="process"
              disabled={
                !files.length ||
                busy
              }
              onClick={run}
            >

              {busy
                ? "Processing..."
                : done
                ? "Process Again"
                : "Process File"}

              <ArrowDownToLine
                size={19}
              />

            </button>

            {/* SUCCESS */}

            {done && (

              <div className="success">

                <Check />

                Your file was created
                and downloaded successfully.

              </div>

            )}

          </div>

        </div>

      </main>

      <Footer />
    </>
  );
}

/* =========================================================
   ALL TOOLS
========================================================= */

function Tools() {

  return (
    <>
      <Header />

      <main className="section container allTools">

        <span className="eyebrow">
          DESIPDF
        </span>

        <h1>
          All PDF Tools
        </h1>

        <p className="lead">
          Choose the tool you need.
        </p>

        <div className="toolGrid">

          {tools.map(
            (tool) => (
              <ToolCard
                key={tool.id}
                tool={tool}
              />
            )
          )}

        </div>

      </main>

      <Footer />
    </>
  );
}
/* =========================================================
   LOADING SCREEN
========================================================= */

function LoadingScreen() {
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    const timer = setTimeout(() => {
      setLoading(false);
    }, 1800);

    return () => clearTimeout(timer);
  }, []);

  if (!loading) return null;

  return (
    <div className="loadingScreen">
      <div className="loadingContent">

        <div className="loadingLogo">
          <div className="pdfShape">
            <span>PDF</span>
          </div>
        </div>

        <h1>
          Desi<span>PDF</span>
        </h1>

        <p>Simple. Fast. Free.</p>

        <div className="loadingBar">
          <div className="loadingProgress"></div>
        </div>

      </div>
    </div>
  );
}
/* =========================================================
   APP
========================================================= */
function App() {
  return (
    <>
      <LoadingScreen />

      <Routes>

        <Route
          path="/"
          element={<Home />}
        />

        <Route
          path="/tools"
          element={<Tools />}
        />

        <Route
          path="/tool/:id"
          element={<ToolPage />}
        />

        <Route
          path="*"
          element={<Home />}
        />

      </Routes>
    </>
  );
}


/* =========================================================
   START APP
========================================================= */

createRoot(
  document.getElementById(
    "root"
  )
).render(
  <BrowserRouter>
    <App />
  </BrowserRouter>
);
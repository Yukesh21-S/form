import { encryptPDF } from "@pdfsmaller/pdf-encrypt";
import { execFile } from "child_process";
import { mkdtemp, readFile, rm, writeFile } from "fs/promises";
import os from "os";
import path from "path";
import { promisify } from "util";
import { GET as generatePpt } from "../generate-ppt/route";

export const runtime = "nodejs";

const execFileAsync = promisify(execFile);

function resolvePdfPassword(req: Request) {
  const url = new URL(req.url);
  return (
    url.searchParams.get("password")?.trim() ||
    process.env.PDF_PASSWORD?.trim() ||
    ""
  );
}

async function encryptPdfWithPdfLib(pdfBuffer: Buffer, password: string) {
  const encrypted = await encryptPDF(new Uint8Array(pdfBuffer), password, {
    ownerPassword: password,
    algorithm: "AES-256",
  });
  return Buffer.from(encrypted);
}

async function convertWithLibreOffice(pptxPath: string, tmpDir: string) {
  const sofficePath = process.env.SOFFICE_PATH?.trim();
  const sofficeCommands = [
    ...(sofficePath ? [sofficePath] : []),
    "soffice",
    "soffice.exe",
  ];
  let convertError: unknown = null;

  for (const command of sofficeCommands) {
    try {
      await execFileAsync(command, [
        "--headless",
        "--nologo",
        "--nodefault",
        "--norestore",
        "--nolockcheck",
        "--invisible",
        "--convert-to",
        "pdf:impress_pdf_Export",
        "--outdir",
        tmpDir,
        pptxPath,
      ]);
      convertError = null;
      break;
    } catch (error) {
      convertError = error;
    }
  }

  if (convertError) {
    throw convertError;
  }
}

async function convertPptxBufferToPdf(pptxBuffer: Buffer) {
  const tmpDir = await mkdtemp(path.join(os.tmpdir(), "report-pdf-"));
  const pptxPath = path.join(tmpDir, "report.pptx");

  try {
    await writeFile(pptxPath, pptxBuffer);
    await convertWithLibreOffice(pptxPath, tmpDir);
    const pdfPath = path.join(tmpDir, "report.pdf");
    return await readFile(pdfPath);
  } finally {
    await rm(tmpDir, { recursive: true, force: true });
  }
}

export async function GET(req: Request) {
  const pptResponse = await generatePpt(req as any);

  if (!pptResponse.ok) {
    return pptResponse;
  }

  const url = new URL(req.url);
  const formId = url.searchParams.get("formId") || "report";
  const password = resolvePdfPassword(req);

  try {
    const pptArrayBuffer = await pptResponse.arrayBuffer();
    const pptBuffer = Buffer.from(pptArrayBuffer);
    let pdfBuffer = await convertPptxBufferToPdf(pptBuffer);

    if (password) {
      pdfBuffer = await encryptPdfWithPdfLib(pdfBuffer, password);
    }

    return new Response(pdfBuffer, {
      headers: {
        "Content-Type": "application/pdf",
        "Content-Disposition": `attachment; filename=report-${formId}.pdf`,
      },
    });
  } catch (error) {
    console.error("PDF generation failed:", error);
    return Response.json(
      {
        error:
          "PDF generation failed. Ensure LibreOffice is installed or set SOFFICE_PATH. For password protection: set PDF_PASSWORD.",
      },
      { status: 500 }
    );
  }
}
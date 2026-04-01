import chromium from "@sparticuz/chromium";
import puppeteer from "puppeteer-core";

const PAGE_SIZE = {
  width: "215.9mm",
  height: "330.2mm",
};

const PDF_TARGETS = {
  msc: {
    printPath: (id) => `/laporan/msc/cetak/${encodeURIComponent(id)}`,
    fileName: (id) => `laporan-msc-${id}.pdf`,
  },
  "cushion-gum": {
    printPath: (id) => `/laporan/cushion-gum/cetak/${encodeURIComponent(id)}`,
    fileName: (id) => `laporan-cushion-gum-${id}.pdf`,
  },
  "gum-cord": {
    printPath: (id) => `/laporan/gum-cord/cetak/${encodeURIComponent(id)}`,
    fileName: (id) => `laporan-gum-cord-${id}.pdf`,
  },
};

function getOrigin(req) {
  const proto = req.headers["x-forwarded-proto"] || "https";
  const host = req.headers["x-forwarded-host"] || req.headers.host;
  return `${proto}://${host}`;
}

async function launchBrowser() {
  const executablePath = await chromium.executablePath();
  return puppeteer.launch({
    args: chromium.args,
    defaultViewport: chromium.defaultViewport,
    executablePath,
    headless: chromium.headless,
  });
}

export default async function handler(req, res) {
  const { kind, id } = req.query || {};
  const target = PDF_TARGETS[kind];

  if (!target || !id) {
    return res.status(400).json({ error: "Parameter kind/id tidak valid." });
  }

  const browser = await launchBrowser();

  try {
    const page = await browser.newPage();
    const cookieHeader = req.headers.cookie;
    if (cookieHeader) {
      await page.setExtraHTTPHeaders({ cookie: cookieHeader });
    }

    const url = `${getOrigin(req)}${target.printPath(id)}`;
    await page.goto(url, {
      waitUntil: "networkidle0",
      timeout: 90000,
    });

    await page.addStyleTag({
      content: `
        .print-toolbar { display: none !important; }
        body { background: #ffffff !important; }
        .sheet { margin: 0 auto !important; }
      `,
    });
    await page.emulateMediaType("print");

    const pdf = await page.pdf({
      printBackground: true,
      preferCSSPageSize: true,
      width: PAGE_SIZE.width,
      height: PAGE_SIZE.height,
      margin: {
        top: "0mm",
        right: "0mm",
        bottom: "0mm",
        left: "0mm",
      },
    });

    res.setHeader("Content-Type", "application/pdf");
    res.setHeader(
      "Content-Disposition",
      `attachment; filename="${target.fileName(id)}"`
    );
    return res.status(200).send(Buffer.from(pdf));
  } catch (error) {
    console.error("PDF render failed:", error);
    return res.status(500).json({ error: "Gagal membuat PDF." });
  } finally {
    await browser.close();
  }
}

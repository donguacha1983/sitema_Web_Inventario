const ExcelJS = require('exceljs');
const path = require('path');

const FILE_PATH = path.join(__dirname, 'data', 'inventario.xlsx');
const RETRYABLE_CODES = new Set(['EBUSY', 'EPERM', 'UNKNOWN']);
const MAX_RETRIES = 6;
const RETRY_DELAY_MS = 300;

let excelIoQueue = Promise.resolve();

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function enqueueExcelIo(task) {
  const run = excelIoQueue.then(() => task());
  excelIoQueue = run.catch(() => undefined);
  return run;
}

async function runWithRetry(task) {
  let lastError;

  for (let attempt = 1; attempt <= MAX_RETRIES; attempt += 1) {
    try {
      return await task();
    } catch (error) {
      lastError = error;

      if (!RETRYABLE_CODES.has(error.code) || attempt === MAX_RETRIES) {
        throw error;
      }

      await sleep(RETRY_DELAY_MS);
    }
  }

  throw lastError;
}

// Carga el archivo Excel principal; si no existe, lo crea.
async function getWorkbook() {
  return enqueueExcelIo(async () => {
    const workbook = new ExcelJS.Workbook();

    try {
      await runWithRetry(() => workbook.xlsx.readFile(FILE_PATH));
      return workbook;
    } catch (error) {
      if (error.code === 'ENOENT') {
        await runWithRetry(() => workbook.xlsx.writeFile(FILE_PATH));
        return workbook;
      }

      if (RETRYABLE_CODES.has(error.code)) {
        throw new Error('El archivo Excel está en uso. Cierre inventario.xlsx en Excel y vuelva a intentar.');
      }

      throw error;
    }
  });
}

// Guarda en disco el libro de trabajo actual.
async function saveWorkbook(workbook) {
  await enqueueExcelIo(async () => {
    try {
      await runWithRetry(() => workbook.xlsx.writeFile(FILE_PATH));
    } catch (error) {
      if (RETRYABLE_CODES.has(error.code)) {
        throw new Error('No se pudo guardar: inventario.xlsx está bloqueado por otro proceso.');
      }

      throw error;
    }
  });
}

module.exports = { getWorkbook, saveWorkbook };

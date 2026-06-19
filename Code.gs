/**
 * Code.gs
 * =======
 * Entry point utama untuk Web App Dashboard Kominfo.
 * Menangani HTTP GET request dan menyajikan halaman HTML.
 */

/**
 * Handler untuk HTTP GET request.
 * Membuat dan mengembalikan halaman web dari template 'index'.
 *
 * @param {Object} e - Event parameter dari request
 * @returns {GoogleAppsScript.HTML.HtmlOutput} Halaman HTML yang di-render
 */
function doGet(e) {
  const template = HtmlService.createTemplateFromFile('index');

  return template.evaluate()
    .setTitle('Dashboard Kominfo - Content Planning')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

/**
 * Menyertakan file HTML eksternal (CSS/JS) ke dalam template utama.
 * Digunakan dengan scriptlet <?!= include('filename') ?> di dalam file HTML.
 *
 * @param {string} filename - Nama file HTML yang akan di-include (tanpa ekstensi .html)
 * @returns {string} Konten HTML dari file yang di-include
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

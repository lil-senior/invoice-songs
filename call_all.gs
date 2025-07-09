/**
 * Runs all purchase extraction scripts in sequence.
 * Useful as a single entry point for triggering all workflows.
 */
function run_all_scripts() {
  extractBandcampPurchases();
  extractBeatportPurchases();
  extractApplePurchases();
  extractJunoDownloadPurchases();
}

/* Other indexes and funds to compare against: month-end levels of a total-return index
   (dividends reinvested), 100 at the first month stored. Filled in and kept current by
   update_spx.py, which runs on the site this app is served from. Until its first run the
   list is empty and the app compares against the S&P 500 only. */
const BENCH = {"v":1,"updated":null,"series":[]};
if (typeof module !== 'undefined') module.exports = { BENCH };

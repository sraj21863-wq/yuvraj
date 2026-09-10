/* Where the dashboard looks for live data, and how old data may get before it
   stops being called live.

   endpoint      the Stage 2 data service (pipeline/serve.py). Leave it as an
                 empty string to run the dashboard purely off the data.js
                 snapshot -- the page then says "SNAPSHOT" rather than "LIVE",
                 which is the point.
   warnAfterHours / staleAfterHours
                 how long after the pipeline last ran the badge turns amber and
                 then red. Booking is looked at through the working day, so a
                 payload older than one working day is not a live dashboard. */
window.OB_CONFIG = {
  endpoint: 'http://127.0.0.1:8787/data.json',
  warnAfterHours: 6,
  staleAfterHours: 30,
  title: 'Order Booking Performance',
  currency: '₹ Cr'
};

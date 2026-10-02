/* ============================================================
   fa-config.js : Fixed Asset module configuration
   ============================================================ */

const FA_CONFIG = {

  moduleCode : "FA",
  moduleName : "Fixed Asset Number Request",
  appVersion : "v4.0.0",

  /* ---- Power Automate HTTP trigger URLs ----
     Paste each URL copied from its flow trigger.
     Plain URL only, inside the quotes.            */

  requestOtpUrl : "https://ca5fc5190790e573a9eafd8b611366.91.environment.api.powerplatform.com:443/powerautomate/automations/direct/cu/11/workflows/33f0420df9d24fa581047ade11d2030b/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=2oMHnIsXj9_3fZjUGbxUx8Urya24YvMGQsZgeVdiMZ4",

  verifyOtpUrl  : "https://ca5fc5190790e573a9eafd8b611366.91.environment.api.powerplatform.com:443/powerautomate/automations/direct/cu/04/workflows/f6be59a13d4945efb46def390e9c78ca/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=z9sOqNYNyFxnlBceUtloiMg0xCEHqFTh5EX1Pm3RYGQ",

  masterDataUrl : "https://ca5fc5190790e573a9eafd8b611366.91.environment.api.powerplatform.com:443/powerautomate/automations/direct/cu/08/workflows/7322454bc1fd4672acc2d8d823bf7325/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=OjAR23Z_cgX6YS_QbBw-VGBm70xtWpBmKUi63LwAQ-4",

  submitUrl     : "https://ca5fc5190790e573a9eafd8b611366.91.environment.api.powerplatform.com:443/powerautomate/automations/direct/cu/17/workflows/939c0c03bb19420e93f498d9d393256e/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=MSHOZQZffVGR0ltB70eWbDtQe4yNHejUCLAm87zwqSQ",

  myRequestsUrl : "https://ca5fc5190790e573a9eafd8b611366.91.environment.api.powerplatform.com:443/powerautomate/automations/direct/cu/02/workflows/cde80967fbc04fd99133353ed8f7a576/triggers/manual/paths/invoke?api-version=1",

  /* ---- Access control ---- */

  allowedDomain : "cgc.com.my",

  /* ---- OTP behaviour (must match the flows) ---- */

  otpLength       : 6,
  otpValidSeconds : 60,
  resendCooldown  : 60,
  sessionMinutes  : 30,
  sessionKey      : "fa_session",

  /* ---- Form limits ---- */

  maxAmount : 100000000

};

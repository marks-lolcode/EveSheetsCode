function debugJaniceOne() {
  const apiKey = getRequiredScriptProperty_("JANICE_API_KEY");
  const log = makeLogger_("JaniceDebug", new Date().toISOString());
  const response = janicePricer_(apiKey, 2, ["Platinum Technite"], log);
  log.info(JSON.stringify(response, null, 2));
}
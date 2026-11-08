/**
 * Configuration helpers.
 *
 * Secrets and environment-specific IDs live in Script Properties
 * (Project Settings > Script properties), never in the source code.
 * The expected keys for this project are listed in config.example.json.
 */

/** Returns a Script Property or the default value when it is not set. */
function getProperty_(key, defaultValue) {
  const value = PropertiesService.getScriptProperties().getProperty(key);
  return (value === null || value === '') ? (defaultValue === undefined ? null : defaultValue) : value;
}

/** Returns a Script Property or throws a clear error when it is missing. */
function getRequiredProperty_(key) {
  const value = getProperty_(key, null);
  if (value === null) throw new Error('Missing Script Property: ' + key + ' (see config.example.json)');
  return value;
}

/** Returns a Script Property parsed as JSON, or the default value when missing or invalid. */
function getJsonProperty_(key, defaultValue) {
  const raw = getProperty_(key, null);
  if (raw === null) return defaultValue;
  try {
    return JSON.parse(raw);
  } catch (err) {
    console.error({ severity: 'ERROR', message: 'Script Property ' + key + ' is not valid JSON', error: err.message });
    return defaultValue;
  }
}

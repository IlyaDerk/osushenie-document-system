/** Shared compatibility helpers for the main and standalone Web App projects. */
function documentArchitectureNormalizeText_(value) {
  return String(value == null ? '' : value)
    .replace(/\u00a0/g, ' ')
    .trim()
    .replace(/\s+/g, ' ');
}

/**
 * Extracts a positive integer only from a trailing «№N» instance suffix.
 * An explicitly populated v2 number always disables legacy fallback.
 */
function documentArchitectureExtractLegacyNumber_(currentType, canonicalType,
  existingNumber) {
  if (documentArchitectureNormalizeText_(existingNumber) !== '') return '';
  const current = documentArchitectureNormalizeText_(currentType);
  const canonical = documentArchitectureNormalizeText_(canonicalType);
  if (!current || !canonical || current === canonical) return '';
  const escaped = canonical.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const match = current.match(new RegExp(
    '^' + escaped + '\\s*(?:[\u2014\u2013-]\\s*)?№\\s*(\\d+)$', 'i'
  ));
  if (!match) return '';
  const number = Number(match[1]);
  return Number.isSafeInteger(number) && number > 0 ? String(number) : '';
}

/** Detects a legacy-looking suffix that cannot be migrated safely. */
function documentArchitectureHasLegacyMarker_(currentType, canonicalType) {
  const current = documentArchitectureNormalizeText_(currentType);
  const canonical = documentArchitectureNormalizeText_(canonicalType);
  if (!current || !canonical) return false;
  const escaped = canonical.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  return new RegExp('^' + escaped + '\\s*(?:[\u2014\u2013-]\\s*)?№', 'i')
    .test(current);
}

/** Shared active-content policy: applies before approval and again at the native boundary. */
const activeFunction =
  /\b(WEBSERVICE|STOCKHISTORY|IMPORTDATA|IMPORTXML|IMPORTHTML|IMPORTRANGE|IMPORTFEED|RTD|CALL|EVALUATE|EXEC|FILES|FOPEN|FWRITE|FREAD|IMAGE)\s*\(/i;
// Bracketed references outside string literals are not admitted yet. This rejects external
// workbook indices/aliases regardless of quoted-sheet punctuation; structured table references
// also require a future explicit footprint contract.
const externalWorkbook = /\[|\]/;
function formulaTokens(value: string): string {
  return value.replace(/"(?:[^"]|"")*"/g, '""');
}
export function isUnsafeFormula(value: string): boolean {
  return (
    activeFunction.test(value) || value.includes('|') || externalWorkbook.test(formulaTokens(value))
  );
}

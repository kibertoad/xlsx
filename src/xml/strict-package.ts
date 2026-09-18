// ISO 29500 strict detection, shared by every reader that resolves elements or
// relationships by transitional QName. A strict package arrives as a plain
// `.xlsx` and differs only in its namespaces, so without this the mismatch
// surfaces as a missing relationship, an unexpected root element, or a workbook
// that loads with no sheets in it.

import { OpenXmlNotImplementedError } from '../utils/exceptions.js';
import { parseQName, STRICT_NS_ROOT } from './namespaces.js';

const strictPackageError = (found: string): OpenXmlNotImplementedError =>
  new OpenXmlNotImplementedError(
    'ISO 29500 strict ("Strict Open XML Spreadsheet" in Excel\'s Save As dialog) is not supported:' +
      ` ${found}. Re-save the file from Excel as "Excel Workbook (.xlsx)" to get the transitional` +
      ' form.',
  );

/**
 * Throw when `rootName`, a Clark-notation element name, is in the strict
 * namespace family. Belongs on a reader's root-element mismatch path: the
 * transitional lookup has already failed by then, and strict is the one cause
 * worth naming rather than reporting as an unexpected root.
 */
export function assertNotStrictRoot(rootName: string): void {
  if (!parseQName(rootName).ns.startsWith(STRICT_NS_ROOT)) return;
  throw strictPackageError(`part root element is "${rootName}"`);
}

/**
 * Throw when any of `types` is the strict spelling of an OPC relationship type.
 * Strict renames the relationship namespace but not the package namespace, so
 * the container still parses and only its officeDocument pointer goes
 * unrecognised.
 */
export function assertNotStrictRelTypes(types: Iterable<string>): void {
  for (const type of types) {
    if (type.startsWith(STRICT_NS_ROOT)) {
      throw strictPackageError(`package relationship type is "${type}"`);
    }
  }
}

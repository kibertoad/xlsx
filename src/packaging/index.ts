// Phase 1 §6 packaging entrypoint. Manifest + Relationships are the
// two structural files every OOXML zip carries; doc properties (core,
// app, custom) follow in the next bootstrap-style turn.

export type { CoreProperties } from './core.js';
export {
  corePropsFromBytes,
  corePropsToBytes,
  makeCoreProperties,
  setWorkbookCategory,
  setWorkbookCreator,
  setWorkbookDescription,
  setWorkbookKeywords,
  setWorkbookLastModifiedBy,
  setWorkbookSubject,
  setWorkbookTitle,
} from './core.js';
export type { CustomProperties, CustomProperty } from './custom.js';
export {
  appendCustomProperty,
  customPropsFromBytes,
  customPropsToBytes,
  findCustomPropertyByName,
  getCustomPropertyValue,
  listCustomProperties,
  makeAsciiStringValue,
  makeBoolValue,
  makeCustomProperties,
  makeDateValue,
  makeDoubleValue,
  makeFiletimeValue,
  makeIntValue,
  makeStringValue,
  readBoolValue,
  readDoubleValue,
  readFiletimeValue,
  readIntValue,
  readStringValue,
  removeCustomProperty,
  setCustomBoolProperty,
  setCustomDateProperty,
  setCustomNumberProperty,
  setCustomStringProperty,
} from './custom.js';
export type { ExtendedProperties } from './extended.js';
export {
  extendedPropsFromBytes,
  extendedPropsToBytes,
  makeExtendedProperties,
  setWorkbookAppVersion,
  setWorkbookApplication,
  setWorkbookCompany,
  setWorkbookHyperlinkBase,
  setWorkbookManager,
} from './extended.js';
export type { DefaultEntry, Manifest, OverrideEntry } from './manifest.js';
export {
  addDefault,
  addOverride,
  findOverride,
  findOverrideByContentType,
  makeManifest,
  manifestFromBytes,
  manifestToBytes,
} from './manifest.js';
export type { Relationship, Relationships } from './relationships.js';
export {
  appendRel,
  findAllByType,
  findById,
  findByType,
  makeRelationships,
  relsFromBytes,
  relsToBytes,
} from './relationships.js';

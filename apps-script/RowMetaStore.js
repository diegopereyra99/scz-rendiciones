
function setRowMeta_(rowIndex, metaObj) {
  const props = PropertiesService.getDocumentProperties();
  props.setProperty(PROP_ROW_META_PREFIX + String(rowIndex), JSON.stringify(metaObj || {}));
}

function getRowMeta_(rowIndex) {
  const props = PropertiesService.getDocumentProperties();
  const raw = props.getProperty(PROP_ROW_META_PREFIX + String(rowIndex));
  if (!raw) return {};
  try { return JSON.parse(raw); } catch (e) { return {}; }
}

function deleteRowMetaRange_(startRow, count) {
  const props = PropertiesService.getDocumentProperties();
  for (let i = 0; i < count; i++) {
    props.deleteProperty(PROP_ROW_META_PREFIX + String(startRow + i));
  }
}

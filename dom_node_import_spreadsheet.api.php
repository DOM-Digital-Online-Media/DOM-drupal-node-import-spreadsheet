<?php

/**
 * @file
 * Describe hooks provided by the DOM node import from spreadsheet module.
 */

/**
 * @addtogroup hooks
 * @{
 */

/**
 * Alter translated data, before save node.
 *
 * @param array $context array keys:
 *   - entity_key: translated entity key (entity_type:entity_id)
 *   - node: original node
 *   - spreadsheet: imported spreadsheet
 *   - settings: submitted form data
 */
function hook_dom_node_import_spreadsheet_save_translate_alter(&$data, $context) {
  // skip hungarian translation
  if ($context['langcode'] == 'hu') {
    $data = [];
  }
}

/**
 * Alter imported data, before create and save node.
 *
 * @param array $data
 *
 * @param array $context array keys:
 *   - node: original node
 *   - spreadsheet: imported spreadsheet
 *   - settings: submitted form data
 */
function hook_dom_node_import_spreadsheet_import_alter(&$data, $context) {
  // Change original node title.
  if ($data['original']['title'] == 'Foobar') {
    $data['original']['title'] = 'Test node';
  }
}

/**
 * Alter exported data, before write to spreadsheet.
 *
 * @param array $result
 * @param array $context array keys:
 *   - languages: languages object array
 *   - spreadsheet: exported spreadsheet
 *   - node: exported node
 */
function hook_dom_node_import_spreadsheet_write_spreadsheet_alter(&$result, $context) {
  // skip row, if key contain '[entity]' string
  foreach ($result as $index => $value) {
    if (strpos($index, '[entity]') !== FALSE) {
      unset($result[$index]);
    }
  }
}

/**
 * @} End of "addtogroup hooks".
 */
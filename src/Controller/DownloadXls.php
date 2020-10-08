<?php

namespace Drupal\dom_node_import_spreadsheet\Controller;

use Drupal\Component\Utility\Unicode;
use Drupal\Core\Controller\ControllerBase;
use Drupal\Core\Entity\ContentEntityInterface;
use Drupal\Core\Entity\EntityManagerInterface;
use Drupal\Core\Extension\ModuleHandlerInterface;
use Drupal\Core\Field\FieldItemListInterface;
use Drupal\Core\File\FileSystemInterface;
use Drupal\Core\Language\LanguageInterface;
use Drupal\Core\Render\Element;
use Drupal\field\Entity\FieldConfig;
use Drupal\node\Entity\Node;
use Drupal\tmgmt\Data;
use Drupal\tmgmt\SourceManager;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Protection;
use Symfony\Component\DependencyInjection\ContainerInterface;
use Symfony\Component\HttpFoundation\BinaryFileResponse;
use Symfony\Component\HttpFoundation\File\MimeType\MimeTypeGuesserInterface;
use Symfony\Component\HttpFoundation\Response;

/**
 * Returns responses for DOM node import from spreadsheet routes.
 */
class DownloadXls extends ControllerBase {

  /**
   * @var \Drupal\Core\File\FileSystemInterface $file_system
   */
  protected $file_system;

  /**
   * @var \Drupal\tmgmt\SourceManager $sourceManager
   */
  protected $sourceManager;

  /**
   * @var \Drupal\tmgmt\Data $tmgmt_data
   */
  protected $tmgmt_data;

  /**
   * @var \PhpOffice\PhpSpreadsheet\Spreadsheet $spreadsheet
   */
  protected $spreadsheet;

  /**
   * @var \Symfony\Component\HttpFoundation\File\MimeType\MimeTypeGuesserInterface $mime_type_guesser
   */
  protected $mime_type_guesser;

  /**
   * {@inheritdoc}
   */
  public function __construct(FileSystemInterface $file_system,
                              SourceManager $sourceManager,
                              Data $tmgmt_data,
                              MimeTypeGuesserInterface $mime_type_guesser,
                              ModuleHandlerInterface $module_handler,
                              EntityManagerInterface $entity_manager) {
    $this->file_system = $file_system;
    $this->sourceManager = $sourceManager;
    $this->tmgmt_data = $tmgmt_data;
    $this->mime_type_guesser = $mime_type_guesser;
    $this->moduleHandler = $module_handler;
    $this->entityManager = $entity_manager;
  }

  /**
   * {@inheritdoc}
   */
  public static function create(ContainerInterface $container) {
    return new static(
      $container->get('file_system'),
      $container->get('plugin.manager.tmgmt.source'),
      $container->get('tmgmt.data'),
      $container->get('file.mime_type.guesser'),
      $container->get('module_handler'),
      $container->get('entity.manager')
    );
  }

  /**
   * Builds the response.
   */
  public function build(Node $node) {
    $required_fields = $this->moduleHandler->invokeAll('dom_drupal_node_import_spreadsheet_required_fields', [$node]);
    $translatable = [];
    // Get languages.
    $languages = $this->languageManager()->getLanguages();
    $default_language = $this->languageManager()->getDefaultLanguage();
    // Get plugin.
    /** @var \Drupal\tmgmt_content\Plugin\tmgmt\Source\ContentEntitySource $plugin */
    $plugin = $this->sourceManager->createInstance('content');
    $result = [];
    foreach ($languages as $language) {
      // Get translated (or original untranslated) entity.
      if ($node->hasTranslation($language->getId())) {
        $translated_node = $node->getTranslation($language->getId());
      }
      else {
        $translated_node = $node->getTranslation(LanguageInterface::LANGCODE_DEFAULT);
      }
      $translatable['node:' . $translated_node->id()][$language->getId()] = $plugin->extractTranslatableData($translated_node);
      $reference_fields = DownloadXls::getReferenceFields($node);
      foreach ($reference_fields as $field_name => $field_definition) {
        $field = $node->get($field_name);
        if (isset($translatable['node:' . $node->id()][$language->getId()][$field_name]) && $field_definition instanceof FieldConfig) {
          foreach (Element::children($translatable['node:' . $node->id()][$language->getId()][$field_name]) as $delta) {
            $field_item = $translatable['node:' . $node->id()][$language->getId()][$field_name][$delta];
            foreach (Element::children($field_item) as $property) {
              if ($target_entity = $this->findReferencedEntity($field, $field_item, $delta, $property)) {
                // Get translated (or original untranslated) entity.
                if ($target_entity->hasTranslation($language->getId())) {
                  $translated_target_entity = $target_entity->getTranslation($language->getId());
                }
                else {
                  $translated_target_entity = $target_entity->getTranslation(LanguageInterface::LANGCODE_DEFAULT);
                }
                $entity_type = $target_entity->getEntityTypeId();
                $translatable[$entity_type . ':' . $target_entity->id()][$language->getId()] = $plugin->extractTranslatableData($translated_target_entity);
              }
            }
          }
        }
        $reference_entities = $field->referencedEntities();
        if (!empty($reference_entities)) {
          // @todo: Make it work with multiple fields.
          $reference_entity = array_shift($reference_entities);
          if ($reference_entity->hasTranslation($language->getId())) {
            $translated_target_entity = $reference_entity->getTranslation($language->getId());
          }
          else {
            $translated_target_entity = $reference_entity->getTranslation(LanguageInterface::LANGCODE_DEFAULT);
          }
          $entity_type = $reference_entity->getEntityTypeId();
          $entity_info = $this->entityManager->getDefinition($entity_type);
          $label = $entity_info->getKey('label');
          $translatable_data = $plugin->extractTranslatableData($translated_target_entity);

          // Put field name as key for reference fields, if this field is exist.
          if (isset($translatable_data[$label])) {
            $translatable['node:' . $node->id()][$language->getId()][$field_name] = $translatable_data[$label];
          }
        }
      }

      foreach ($this->tmgmt_data->flatten($translatable) as $index => $item) {
        list ($data_entity_id, $data_langcode, $data_field_id, $data_other) = explode('][', $index, 4);
        if ((empty($required_fields) || in_array($data_field_id, $required_fields)) && strpos($data_other, 'format') === FALSE) {

          // Split link field, to url and title fields.
          if ($data_other === '0][uri') {
            $data_field_id .= '_uri';
          }
          $result[$data_entity_id][$data_field_id][$data_langcode] = [
            'item' => $item,
            'data_index' => $index,
            'data_other' => $data_other,
          ];
        }
      }
    }

    // Create empty spreadsheet and write data.
    $this->spreadsheet = new Spreadsheet();
    $this->spreadsheet
      ->getProperties()
      ->setCreated(\Drupal::time()->getRequestTime());
    $this->writeXlsx($default_language, $languages, $result, $node);
    $filename = $this->file_system->tempnam('temporary://', 'dnis_') . '.xlsx';
    $objWriter = IOFactory::createWriter($this->spreadsheet, 'Xlsx');
    ob_start();
    $objWriter->save('php://output');
    $output = ob_get_clean();
    $file = file_save_data($output, $filename, FILE_EXISTS_RENAME);
    $file->setTemporary();
    $file->save();
    // Send file to browser.
    $filename = sprintf('dom_node_import_spreadsheet__%d.xlsx', $node->id());
    $mime = $this->mime_type_guesser->guess($file->getFileUri());
    $headers = [
      'Content-Type' => $mime . '; name="' . Unicode::mimeHeaderEncode(basename($file->getFileUri())) . '"',
      'Content-Length' => filesize($file->getFileUri()),
      'Content-Disposition' => 'attachment; filename="' . Unicode::mimeHeaderEncode($filename) . '"',
      'Cache-Control' => 'private',
    ];
    return new BinaryFileResponse($file->getFileUri(), Response::HTTP_OK, $headers, FALSE);
  }

  /**
   * Write data to spreadsheet.
   */
  private function writeXlsx($default_language, array $languages, array $data, Node $node) {
    unset($languages[$default_language->getId()]);
    // Add node ID for first row.
    $this->spreadsheet
      ->getActiveSheet()
      ->setCellValue('A1', $node->id());
    // Add header.
    $header = [
      'A' => 'Entity',
      'B' => 'Key',
      'C' => 'Original text',
    ];
    $column_index = 'C';
    foreach ($languages as $langcode => $language) {
      $header[++$column_index] = mb_strtoupper($langcode);
    }
    foreach ($header as $col => $text) {
      $this->spreadsheet
        ->getActiveSheet()
        ->setCellValue($col . '2', $text);
    }
    $this->spreadsheet
      ->getActiveSheet()
      ->freezePane('C3');
    $row = 3;
    // Data processing.
    foreach ($data as $data_entity_id => $temporary_data) {
      foreach ($temporary_data as $field_name => $field_data) {
        // Entity.
        $coordinate = 'A' . $row;
        $this->spreadsheet
          ->getActiveSheet()
          ->setCellValue($coordinate, $data_entity_id);
        // Key.
        $coordinate = 'B' . $row;
        $this->spreadsheet
          ->getActiveSheet()
          ->setCellValue($coordinate, $field_name);
        // Original text.
        $coordinate = 'C' . $row;
        $text = isset($field_data[$default_language->getId()]['item']['#text']) ? $field_data[$default_language->getId()]['item']['#text'] : '0';
        $this->spreadsheet
          ->getActiveSheet()
          ->calculateColumnWidths()
          ->setCellValue($coordinate, $text);
        // Languages.
        $column_index = 'D';
        foreach ($field_data as $data_langcode => $data_value) {

          // Skip english language.
          if ($data_langcode === 'en') {
            unset($data_value);
          }
          else {
            $coordinate = $column_index++ . $row;
            $text = isset($data_value['item']['#text']) ? $data_value['item']['#text'] : '';
            $text = $text == '1' ? '' : $text;
            $this->spreadsheet
              ->getActiveSheet()
              ->calculateColumnWidths()
              ->setCellValue($coordinate, $text)
              ->getStyle($coordinate)
              ->getProtection()
              ->setLocked(Protection::PROTECTION_UNPROTECTED);
          }
        }
        $row++;
      }
    }
    // Auto size cells.
    foreach (['B', 'C'] as $index) {
      $this->spreadsheet
        ->getActiveSheet()
        ->getColumnDimension($index)
        ->setAutoSize(TRUE);
    }
    // Hide metadata row.
    $this->spreadsheet
      ->getActiveSheet()
      ->getRowDimension(1)
      ->setVisible(FALSE);
    // hide key column and select active cell
//    $this->spreadsheet
//      ->getActiveSheet()
//      ->setSelectedCell('C1')
//      ->getColumnDimension('B')
//      ->setVisible(FALSE);
    $this->spreadsheet->getActiveSheet()->getProtection()->setPassword('FreeBlocking');
    $this->spreadsheet->getActiveSheet()->getProtection()->setSheet(TRUE);
    $this->spreadsheet->setActiveSheetIndex(0);
  }

  /**
   * @return \Drupal\Core\Field\FieldDefinitionInterface[]
   */
  static public function getReferenceFields(ContentEntityInterface $entity) {
    $embeddable_fields = [];
    $field_definitions = $entity->getFieldDefinitions();
    $content_translation_manager = \Drupal::service('content_translation.manager');
    foreach ($field_definitions as $field_name => $field_definition) {
      $storage_definition = $field_definition->getFieldStorageDefinition();
      $property_definitions = $storage_definition->getPropertyDefinitions();
      foreach ($property_definitions as $property_definition) {
        if (in_array($property_definition->getDataType(), ['entity_reference', 'entity_revision_reference']) && $storage_definition->getSetting('target_type') && $content_translation_manager->isEnabled($storage_definition->getSetting('target_type'))) {
          $embeddable_fields[$field_name] = $field_definition;
        }
      }
    }
    return $embeddable_fields;
  }

  /**
   * @return \Drupal\Core\Entity\ContentEntityInterface|null
   */
  protected function findReferencedEntity(FieldItemListInterface $field, array $field_item, $delta, $property) {
    if (isset($field_item[$property]['#id'])) {
      foreach ($field as $item) {
        if ($item->$property instanceof ContentEntityInterface && $item->$property->id() == $field_item[$property]['#id']) {
          return $item->$property;
        }
      }
    }
    elseif ($property == 'target_id' && isset($field_item[$property]['#text'])) {
      $entity_type = $field->getSetting('target_type');
      $controller = $this->entityManager->getStorage($entity_type);
      return $controller->load($field_item[$property]['#text']);
    }
  }

}

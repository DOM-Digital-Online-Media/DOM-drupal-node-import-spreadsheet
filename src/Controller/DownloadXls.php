<?php

namespace Drupal\dom_node_import_spreadsheet\Controller;

use Drupal\Component\Utility\Unicode;
use Drupal\Core\Controller\ControllerBase;
use Drupal\Core\Entity\ContentEntityInterface;
use Drupal\Core\Entity\EntityManagerInterface;
use Drupal\Core\Extension\ModuleHandlerInterface;
use Drupal\Core\Field\FieldItemListInterface;
use Drupal\Core\File\FileSystemInterface;
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
    $translatable = [];
    // get languages
    $languages = $this->languageManager()->getLanguages();
    $default_language = $this->languageManager()->getDefaultLanguage();
    unset($languages[$default_language->getId()]);
    // get plugin
    /** @var \Drupal\tmgmt_content\Plugin\tmgmt\Source\ContentEntitySource $plugin */
    $plugin = $this->sourceManager->createInstance('content');
    $translatable['node:' . $node->id()] = $plugin->extractTranslatableData($node);
    $reference_fields = DownloadXls::getReferenceFields($node);
    foreach ($reference_fields as $field_name => $field_definition) {
      $field = $node->get($field_name);
      if (isset($translatable['node:' . $node->id()][$field_name]) && $field_definition instanceof FieldConfig) {
        foreach (Element::children($translatable['node:' . $node->id()][$field_name]) as $delta) {
          $field_item = $translatable['node:' . $node->id()][$field_name][$delta];
          foreach (Element::children($field_item) as $property) {
            if ($target_entity = $this->findReferencedEntity($field, $field_item, $delta, $property)) {
              $entity_type = $target_entity->getEntityTypeId();
              $translatable[$entity_type . ':' . $target_entity->id()] = $plugin->extractTranslatableData($target_entity);
            }
          }
        }
      }
    }
    $result = [];
    foreach ($this->tmgmt_data->flatten($translatable) as $index => $item) {
      if (isset($item['#translate']) && $item['#translate'] == TRUE) {
        $result[$index] = $item;
      }
    }
    // create empty spreadsheet
    $this->spreadsheet = new Spreadsheet();
    $this->spreadsheet
      ->getProperties()
      ->setCreated(\Drupal::time()->getRequestTime());
    // write spreadsheet data
    $context = [
      'languages' => $languages,
      'spreadsheet' => $this->spreadsheet,
      'node' => $node,
    ];
    // hook_dom_node_import_spreadsheet_write_spreadsheet_alter()
    $this->moduleHandler->alter('dom_node_import_spreadsheet_write_spreadsheet', $result, $context);
    $this->writeXlsx($languages, $result, $node);
    $filename = $this->file_system->tempnam('temporary://', 'dnis_') . '.xlsx';
    $objWriter = IOFactory::createWriter($this->spreadsheet, 'Xlsx');
    ob_start();
    $objWriter->save('php://output');
    $output = ob_get_clean();
    $file = file_save_data($output, $filename, FILE_EXISTS_RENAME);
    $file->setTemporary();
    $file->save();
    // send file to browser
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
   * Write data to spreadsheet
   */
  private function writeXlsx(array $languages, array $data, Node $node) {
    // add node ID for first row
    $this->spreadsheet
      ->getActiveSheet()
      ->setCellValue('A1', $node->id());
    // add header
    $header = [
      'A' => 'Entity',
      'B' => 'Key',
      'C' => 'Parent label',
      'D' => 'Label',
      'E' => 'Original text',
    ];
    $column_index = 'E';
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
    foreach ($data as $key => $value) {
      // entity
      $coordinate = 'A' . $row;
      $field_path = explode('][', $key);
      $this->spreadsheet
        ->getActiveSheet()
        ->setCellValue($coordinate, $field_path[0]);
      // key
      $coordinate = 'B' . $row;
      $this->spreadsheet
        ->getActiveSheet()
        ->setCellValue($coordinate, substr($key, strlen($field_path[0]) + 2));
      // parent label
      $coordinate = 'C' . $row;
      $parent_label = isset($value['#parent_label'][0]) ? $value['#parent_label'][0] : '';
      $this->spreadsheet
        ->getActiveSheet()
        ->calculateColumnWidths()
        ->setCellValue($coordinate, $parent_label);
      // label
      $coordinate = 'D' . $row;
      $label = isset($value['#label']) ? $value['#label'] : '';
      $this->spreadsheet
        ->getActiveSheet()
        ->calculateColumnWidths()
        ->setCellValue($coordinate, $label);
      // text
      $coordinate = 'E' . $row;
      $text = isset($value['#text']) ? $value['#text'] : '';
      $this->spreadsheet
        ->getActiveSheet()
        ->calculateColumnWidths()
        ->setCellValue($coordinate, $text);
      // languages
      $column_index = 'E';
      while (isset($header[++$column_index])) {
        $coordinate = $column_index . $row;
        $text = isset($value['#text']) ? $value['#text'] : '';
        $this->spreadsheet
          ->getActiveSheet()
          ->calculateColumnWidths()
          ->setCellValue($coordinate, $text)
          ->getStyle($coordinate)
          ->getProtection()
          ->setLocked(Protection::PROTECTION_UNPROTECTED);
      }
      // auto size cells
      foreach (['C', 'D', 'E'] as $index) {
        $this->spreadsheet
          ->getActiveSheet()
          ->getColumnDimension($index)
          ->setAutoSize(true);
      }
      $row++;
    }
    // hide metadata row
    $this->spreadsheet
      ->getActiveSheet()
      ->getRowDimension(1)
      ->setVisible(FALSE);
    // hide key column and select active cell
    $this->spreadsheet
      ->getActiveSheet()
      ->setSelectedCell('C1')
      ->getColumnDimension('B')
      ->setVisible(FALSE);
    $this->spreadsheet->getActiveSheet()->getProtection()->setPassword('FreeBlocking');
    $this->spreadsheet->getActiveSheet()->getProtection()->setSheet(true);
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

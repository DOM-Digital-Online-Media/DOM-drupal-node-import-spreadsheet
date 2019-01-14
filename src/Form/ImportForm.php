<?php

namespace Drupal\dom_node_import_spreadsheet\Form;

use Drupal\Core\Entity\EntityTypeBundleInfoInterface;
use Drupal\Core\Entity\EntityTypeManagerInterface;
use Drupal\Core\Extension\ModuleHandlerInterface;
use Drupal\Core\File\FileSystemInterface;
use Drupal\Core\Form\FormBase;
use Drupal\Core\Form\FormStateInterface;
use Drupal\Core\Language\LanguageManagerInterface;
use Drupal\file\Entity\File;
use Drupal\filter\Entity\FilterFormat;
use Drupal\node\Entity\Node;
use PhpOffice\PhpSpreadsheet\IOFactory;
use Symfony\Component\DependencyInjection\ContainerInterface;

/**
 * Import contents from spreadsheet.
 */
class ImportForm extends FormBase {

  /**
   * @var \Drupal\Core\Entity\EntityTypeBundleInfoInterface $entity_bundles
   */
  protected $entity_bundles;

  /**
   * @var \Drupal\Core\File\FileSystemInterface $file_system
   */
  protected $file_system;

  /**
   * @var \Drupal\Core\Language\LanguageManagerInterface $language_manager
   */
  protected $language_manager;

  /**
   * @var \Drupal\Core\Entity\EntityTypeManagerInterface $entity_type_manager
   */
  protected $entity_type_manager;

  /**
   * @var \Drupal\Core\Extension\ModuleHandlerInterface $module_handler
   */
  protected $module_handler;

  /**
   * @var \PhpOffice\PhpSpreadsheet\Spreadsheet $spreadsheet
   */
  protected $spreadsheet;

  /**
   * @var \Drupal\node\Entity\Node $node
   */
  protected $node;

  /**
   * ImportForm constructor.
   */
  public function __construct(EntityTypeBundleInfoInterface $entity_bundles,
                              FileSystemInterface $file_system,
                              LanguageManagerInterface $language_manager,
                              EntityTypeManagerInterface $entity_type_manager,
                              ModuleHandlerInterface $module_handler) {
    $this->entity_bundles = $entity_bundles;
    $this->file_system = $file_system;
    $this->language_manager = $language_manager;
    $this->entity_type_manager = $entity_type_manager;
    $this->module_handler = $module_handler;
  }

  /**
   * {@inheritdoc}
   */
  public static function create(ContainerInterface $container) {
    return new static(
      $container->get('entity_type.bundle.info'),
      $container->get('file_system'),
      $container->get('language_manager'),
      $container->get('entity_type.manager'),
      $container->get('module_handler')
    );
  }

  /**
   * {@inheritdoc}
   */
  public function getFormId() {
    return 'dom_node_import_spreadsheet';
  }

  /**
   * {@inheritdoc}
   */
  public function buildForm(array $form, FormStateInterface $form_state) {
    $form['import'] = [
      '#type' => 'managed_file',
      '#required' => TRUE,
      '#title' => $this->t('Import file'),
      '#description' => $this->t('Allowed formats: XLSX'),
      '#size' => 30,
      '#upload_validators' => [
        'file_validate_extensions' => [
          'xlsx',
        ],
      ],
    ];
    $form['overwrite'] = [
      '#type' => 'checkbox',
      '#title' => $this->t('Overwrite existing translations.'),
    ];
    $form['publish'] = [
      '#type' => 'checkbox',
      '#title' => $this->t('Publish new translations.'),
    ];
    $form['new_node'] = [
      '#type' => 'fieldset',
      '#tree' => FALSE,
      '#title' => $this->t('Only affect, if create new node'),
      '#collapsible' => FALSE,
      '#collapsed' => FALSE,
    ];
    $node_bundles = [];
    foreach ($this->entity_bundles->getBundleInfo('node') as $bundle_id => $bundle) {
      $node_bundles[$bundle_id] = $bundle['label'];
    }
    $form['new_node']['bundle'] = [
      '#type' => 'select',
      '#required' => TRUE,
      '#title' => $this->t('Imported node type'),
      '#options' => $node_bundles,
    ];
    $text_formats = [];
    $filter_format_keys = $this->entity_type_manager->getStorage('filter_format')->getQuery()->execute();
    $filter_formats = FilterFormat::loadMultiple($filter_format_keys);
    foreach ($filter_formats as $key => $format) {
      $text_formats[$key] = $format->label();
    }
    $form['new_node']['text_format'] = [
      '#type' => 'select',
      '#required' => TRUE,
      '#title' => $this->t('Text format'),
      '#options' => $text_formats,
    ];
    $form['submit'] = [
      '#type' => 'submit',
      '#value' => $this->t('Send'),
    ];
    return $form;
  }

  /**
   * {@inheritdoc}
   */
  public function validateForm(array &$form, FormStateInterface $form_state) {
    $fid = $form_state->getValue('import');
    if (!empty($fid)) {
      $fid = array_shift($form_state->getValue('import'));
      $file = File::load($fid);
      $real_path = $this->file_system->realpath($file->getFileUri());
      $this->spreadsheet = IOFactory::createReader('Xlsx')->load($real_path);
      // verify node id
      $nid = (integer) $this->spreadsheet
        ->getActiveSheet()
        ->getCell('A1')
        ->getValue();
      if (!empty($nid)) {
        // upgrade existing node
        $this->node = Node::load($nid);
        if (!$this->node) {
          $form_state->setErrorByName('import', $this->t('Missing node id, loaded: @nid', array('@nid' => $nid)));
        }
      }
    }
  }

  /**
   * {@inheritdoc}
   */
  public function submitForm(array &$form, FormStateInterface $form_state) {
    // get languages
    $languages = $this->language_manager->getLanguages();
    $default_language = $this->language_manager->getDefaultLanguage();
    unset($languages[$default_language->getId()]);
    // get rows and columns
    $rows = $this->spreadsheet
      ->getActiveSheet()
      ->getRowIterator(2);
    $original_data = [];
    $translated_data = [];
    foreach ($rows as $row_index => $row) {
      // element key
      $element_key = $this->spreadsheet
        ->getActiveSheet()
        ->getCell('A' . $row_index)
        ->getValue();
      // original data
      if (!empty($element_key)) {
        $original_data[$element_key] = $this->spreadsheet
          ->getActiveSheet()
          ->getCell('B' . $row_index)
          ->getValue();
        foreach ($row->getCellIterator('B') as $column_index => $cell) {
          $langcode = mb_strtolower($this->spreadsheet
            ->getActiveSheet()
            ->getCell($column_index . '1')
            ->getValue());
          if (!empty($langcode) && array_key_exists($langcode, $languages)) {
            $translated_data[$langcode][$element_key] = $cell->getValue();
          }
        }
      }
    }

    if (empty($this->node)) {
      // new node
      $this->node = Node::create(['type' => $form_state->getValue('bundle')]);
    }
    // hook_dom_node_import_spreadsheet_import_alter(&$data, $context)
    $alterable = [
      'original' => &$original_data,
      'translated' => &$translated_data,
    ];
    $context = [
      'node' => clone $this->node,
      'spreadsheet' => clone $this->spreadsheet,
      'settings' => $form_state->cleanValues()->getValues(),
    ];
    $this->module_handler->alter('dom_node_import_spreadsheet_import', $alterable, $context);
    // save node with source language
    try {
      foreach ($original_data as $element_key => $field_data) {
        if ($this->node->hasField($element_key)) {
          // field exist
          $data = [
            'value' => $field_data,
            'format' => $form_state->getValue('text_format'),
          ];
          $this->node->set($element_key, $data);
        }
      }
      if ($form_state->getValue('publish')) {
        $this->node->set('status', Node::PUBLISHED);
      }
      if (!empty($translated_data)) {
        // save translations
        foreach ($translated_data as $langcode => $lang_data) {
          if (!$this->node->hasTranslation($langcode)) {
            $translated_node = $this->node->addTranslation($langcode);
          }
          elseif ($form_state->getValue('overwrite')) {
            $translated_node = $this->node->getTranslation($langcode);
          }
          else {
            // skip translation
            continue;
          }
          foreach ($lang_data as $element_key => $field_data) {
            if ($this->node->hasField($element_key)) {
              // field exist
              $data = [
                'value' => $field_data,
                'format' => $form_state->getValue('text_format'),
              ];
              $translated_node->set($element_key, $data);
            }
          }
          if ($form_state->getValue('publish')) {
            $translated_node->set('status', Node::PUBLISHED);
          }
        }
      }
      $this->node->save();
      $this->messenger()->addMessage($this->t('%link node successfully imported.', [
        '%link' => $this->node->link('nid: ' . $this->node->id())
      ]));
    }
    catch (\Exception $e) {
      $this->messenger()->addError($this->t('Error, nid: @nid node not translated.', ['@nid' => $this->node->id()]));
    }
  }

}

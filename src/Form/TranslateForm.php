<?php

namespace Drupal\dom_node_import_spreadsheet\Form;

use Drupal\Core\Entity\EntityManagerInterface;
use Drupal\Core\Extension\ModuleHandlerInterface;
use Drupal\Core\File\FileSystemInterface;
use Drupal\Core\Form\FormBase;
use Drupal\Core\Form\FormStateInterface;
use Drupal\Core\Language\LanguageManagerInterface;
use Drupal\Core\Url;
use Drupal\dom_node_import_spreadsheet\Controller\DownloadXls;
use Drupal\file\Entity\File;
use Drupal\node\Entity\Node;
use Drupal\node\NodeInterface;
use Drupal\tmgmt\Data;
use PhpOffice\PhpSpreadsheet\IOFactory;
use Symfony\Component\DependencyInjection\ContainerInterface;

/**
 * Provides a DOM node import from spreadsheet form.
 */
class TranslateForm extends FormBase {

  /**
   * @var \Drupal\Core\File\FileSystemInterface $file_system
   */
  protected $file_system;

  /**
   * @var \Drupal\Core\Language\LanguageManagerInterface $language_manager
   */
  protected $language_manager;

  /**
   * @var \Drupal\tmgmt\Data $tmgmt_data
   */
  protected $tmgmt_data;

  /**
   * @var \Drupal\Core\Extension\ModuleHandlerInterface $module_handler
   */
  protected $module_handler;

  /**
   * @var \PhpOffice\PhpSpreadsheet\Spreadsheet $spreadsheet
   */
  protected $spreadsheet;

  /**
   * @var Node $node
   */
  protected $node;

  /**
   * @var \Drupal\Core\Entity\EntityManagerInterface $entity_manager
   */
  protected $entity_manager;

  /**
   * TranslateForm constructor.
   */
  public function __construct(FileSystemInterface $file_system,
                              LanguageManagerInterface $languageManager,
                              Data $tmgmt_data,
                              ModuleHandlerInterface $module_handler,
                              EntityManagerInterface $entity_manager) {
    $this->file_system = $file_system;
    $this->language_manager = $languageManager;
    $this->tmgmt_data = $tmgmt_data;
    $this->module_handler = $module_handler;
    $this->entity_manager = $entity_manager;
  }

  /**
   * {@inheritdoc}
   */
  public static function create(ContainerInterface $container) {
    return new static(
      $container->get('file_system'),
      $container->get('language_manager'),
      $container->get('tmgmt.data'),
      $container->get('module_handler'),
      $container->get('entity.manager')
    );
  }

  /**
   * {@inheritdoc}
   */
  public function getFormId() {
    return 'dom_node_import_spreadsheet_translate';
  }

  /**
   * {@inheritdoc}
   */
  public function buildForm(array $form, FormStateInterface $form_state, Node $node = NULL) {
    $form['nid'] = [
      '#type' => 'value',
      '#value' => $node->id(),
    ];
    $form['download'] = [
      '#type' => 'link',
      '#title' => $this->t('Download XLS'),
      '#url' => Url::fromRoute('dom_node_import_spreadsheet.translate_form_download_xls', ['node' => $node->id()]),
      '#options' => array(
        'attributes' => array('class' => ['button', 'button-action']),
      ),
    ];
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
    $form['submit'] = [
      '#type' => 'submit',
      '#value' => t('Send'),
    ];
    return $form;
  }

  /**
   * {@inheritdoc}
   */
  public function validateForm(array &$form, FormStateInterface $form_state) {
    $this->node = Node::load($form_state->getValue('nid'));
    $fid = $form_state->getValue('import');
    if (!empty($fid)) {
      $fid = array_shift($form_state->getValue('import'));
      $file = File::load($fid);
      $real_path = $this->file_system->realpath($file->getFileUri());
      $this->spreadsheet = IOFactory::createReader('Xlsx')->load($real_path);
      // verify node id
      $loaded_nid = $this->spreadsheet
        ->getActiveSheet()
        ->getCell('A1')
        ->getValue();
      if ($loaded_nid != $this->node->id()) {
        $form_state->setErrorByName('import', $this->t('Missing node id, loaded: @nid', array('@nid' => $loaded_nid)));
      }
    }
  }

  /**
   * {@inheritdoc}
   */
  public function submitForm(array &$form, FormStateInterface $form_state) {
    $translated_data = [];
    // get languages
    $languages = $this->language_manager->getLanguages();
    $default_language = $this->language_manager->getDefaultLanguage();
    unset($languages[$default_language->getId()]);
    // get rows and columns
    $rows = $this->spreadsheet
      ->getActiveSheet()
      ->getRowIterator(3);
    foreach ($rows as $row_index => $row) {
      // entity key
      $entity_key = $this->spreadsheet
        ->getActiveSheet()
        ->getCell('A' . $row_index)
        ->getValue();
      if (isset($previous_entity_key) && $entity_key != $previous_entity_key) {
        // next entity
        $item = [];
      }
      // element key
      $element_key = $this->spreadsheet
        ->getActiveSheet()
        ->getCell('B' . $row_index)
        ->getValue();
      foreach ($row->getCellIterator('F') as $column_index => $cell) {
        $langcode = mb_strtolower($this->spreadsheet
          ->getActiveSheet()
          ->getCell($column_index . '2')
          ->getValue());
        if (!empty($langcode) && array_key_exists($langcode, $languages)) {
          $item[$langcode][$element_key] = $cell->getValue();
          $translated_data[$entity_key][$langcode] = $this->tmgmt_data->unflatten($item[$langcode]);
        }
      }
      $previous_entity_key = $entity_key;
    }
    // save entity translation
    try {
      foreach ($translated_data as $entity_key => $all_data) {
        // hook_dom_node_import_spreadsheet_save_translate_alter(&$data, $context)
        $context = [
          'entity_key' => $entity_key,
          'node' => clone $this->node,
          'spreadsheet' => clone $this->spreadsheet,
          'settings' => $form_state->cleanValues()->getValues(),
        ];
        list($entity_type, $entity_id) = explode(':', $entity_key);
        $controller = $this->entity_manager->getStorage($entity_type);
        /** @var \Drupal\Core\Entity\ContentEntityInterface $entity */
        $entity = $controller->load($entity_id);
        $this->module_handler->alter('dom_node_import_spreadsheet_save_translate', $all_data, $context);
        foreach ($all_data as $langcode => $data) {
          if (!isset($data) || empty($data)) {
            // skip translation
            continue;
          }
          if (!$entity->hasTranslation($langcode)) {
            $translated_entity = $entity->addTranslation($langcode, $data);
          }
          elseif ($form_state->getValue('overwrite')) {
            $translated_entity = $entity->getTranslation($langcode);
            foreach ($data as $field_name => $field_data) {
              $translated_entity->set($field_name, $field_data);
            }
          }
          if (isset($translated_entity)) {
            if ($form_state->getValue('publish') && $entity instanceof NodeInterface) {
              $translated_entity->set('status', Node::PUBLISHED);
            }
            // copy referenced entities
            foreach (DownloadXls::getReferenceFields($entity) as $field_name => $field_definition) {
              $original_value = $entity->get($field_name)->getValue();
              $translated_entity->set($field_name, $original_value);
            }
            $translated_entity->save();
          }
        }
      }
      $this->messenger()->addMessage($this->t('Node successfully translated.'));
    }
    catch (\Exception $e) {
      $this->messenger()->addError($this->t('Error, node not translated.'));
    }
  }

}

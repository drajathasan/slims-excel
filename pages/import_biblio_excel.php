<?php
/**
 * @author Drajat Hasan
 * @email drajathasan20@gmail.com
 * @create date 2024-09-24 16:49:08
 * @modify date 2024-09-26 23:02:44
 * @desc 
 * - License : GPL-v3
 */

use SLiMS\DB;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx as Reader;

defined('INDEX_AUTH') or die('direct access is not allowed!');

// IP based access limitation
require LIB.'ip_based_access.inc.php';
do_checkIP('smc');
do_checkIP('smc-bibliography');
// start the session
require SB.'admin/default/session.inc.php';
require SIMBIO.'simbio_DB/simbio_dbop.inc.php';
require SIMBIO.'simbio_GUI/table/simbio_table.inc.php';
require SIMBIO.'simbio_GUI/form_maker/simbio_form_table_AJAX.inc.php';
include __DIR__ . '/backup_alert.php';

// privileges checking
$can_read = utility::havePrivilege('bibliography', 'r');
$can_write = utility::havePrivilege('bibliography', 'w');

if (!$can_read) {
  die('<div class="errorBox">'.__('You are not authorized to view this section').'</div>');
}

if (isset($_GET['action']) && $_GET['action'] == 'download_sample') {
    $spreadsheet = new Spreadsheet();
    $spreadsheet->getActiveSheet()
      ->fromArray(
          [[
            'title',
            'gmd_name',
            'edition',
            'isbn_issn',
            'publisher_name',
            'publish_year',
            'collation',
            'series_title',
            'call_number',
            'language_name',
            'place_name',
            'classification',
            'notes',
            'image',
            'sor',
            'authors',
            'topics',
            'item_code'
          ]],  // The data to set
          NULL,        // Array values with this value will not be set
          'A1'         // Top left coordinate of the worksheet range where
                       //    we want to set these values (default is A1)
      );
    $writer = new Xlsx($spreadsheet);
    $tblout ='import_excel_sample';
    header("Content-Type: application/xlsx");
    header("Content-Disposition: attachment; filename=$tblout.xlsx");
    header("Pragma: no-cache");
    $writer->save('php://output');
    exit;
}

if (isset($_POST['doImport'])) {
    $pdo = DB::getInstance();
    $pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);

    try {
        if ($sysconf['index']['type'] == 'index') {
            require MDLBS.'system/biblio_indexer.inc.php';
            // create biblio_indexer class instance
            $indexer = new biblio_indexer($dbs);
        }
    
        $reader = new Reader;
        $reader->setReadDataOnly(true);
        $spreadsheet = $reader->load($_FILES['importFile']['tmp_name']);
        $sheet = $spreadsheet->getSheet($spreadsheet->getFirstSheetIndex());
    
        $state = $pdo->prepare(<<<SQL
        INSERT IGNORE INTO biblio (title, gmd_id, edition,
                    isbn_issn, publisher_id, publish_year,
                    collation, series_title, call_number,
                    language_id, publish_place_id, classification,
                    notes, image, sor, input_date, last_update)
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,now(),now());
        SQL);
    
        $row = 0;
        foreach ($sheet->toArray() as $data) {
            if (($data[0]??'') == 'title') continue;

            foreach ($data as $index => $field) {
                $currentValue = $field;
                switch ($index) {
                    // title formatter
                    case 0:
                      $currentValue = empty($currentValue) ? NULL : $currentValue;
                      break;
          
                    case 1:
                      $currentValue = empty($currentValue) ? NULL : utility::getID($dbs, 'mst_gmd', 'gmd_id', 'gmd_name', $currentValue);
                      break;
          
                    case 4:
                      $currentValue = empty($currentValue) ? NULL : utility::getID($dbs, 'mst_publisher', 'publisher_id', 'publisher_name', $currentValue);
                      break;
          
                    case 9:
                      $currentValue = empty($currentValue) ? NULL : utility::getID($dbs, 'mst_language', 'language_id', 'language_name', $currentValue);
                      break;
          
                    case 10:
                      $currentValue = empty($currentValue) ? NULL : utility::getID($dbs, 'mst_place', 'place_id', 'place_name', $currentValue);
                      break;
                    
                    default:
                      $currentValue = empty($currentValue) ? NULL : $currentValue;
                      break;
                }
    
                $data[$index] = $currentValue;
            }
    
            $authors = $data[15];
            $subjects = $data[16];
            $items = $data[17];
    
            // remove data with tag format
            unset($data[15]);
            unset($data[16]);
            unset($data[17]);

            $state->execute($data);
    
            if ($state) {
                $biblio_id = $pdo->lastInsertId();
      
                if (!empty($authors)) {
                  $biblio_author_sql = $pdo->prepare('INSERT IGNORE INTO biblio_author (biblio_id, author_id, level) VALUES (?,?,?)');
                  $authors = explode('><', trim($authors, '<>'));
                  foreach ($authors as $author) {
                    $author = trim(str_replace(array('>', '<'), '', $author));
                    $author_id = utility::getID($dbs, 'mst_author', 'author_id', 'author_name', $author);
                    $biblio_author_sql->execute([$biblio_id, $author_id, 2]);
                  }
                }
      
                if (!empty($subjects)) {
                  $biblio_subject_sql = $pdo->prepare('INSERT IGNORE INTO biblio_topic (biblio_id, topic_id, level) VALUES (?,?,?)');
                  $subjects = explode('><', trim($subjects, '<>'));
                  foreach ($subjects as $subject) {
                    $subject = trim(str_replace(array('>', '<'), '', $subject));
                    $subject_id = utility::getID($dbs, 'mst_topic', 'topic_id', 'topic', $subject);
                    $biblio_subject_sql->execute([$biblio_id, $subject_id , 2]);
                  }
                }
      
                // items
                if (!empty($items)) {
                  $item_sql = $pdo->prepare('INSERT IGNORE INTO item (biblio_id, item_code) VALUES (?,?)');
                  $item_array = explode('><', $items);
                  foreach ($item_array as $item) {
                    $item = trim(str_replace(array('>', '<'), '', $item));
                    $item_sql->execute([$biblio_id, $item]);
                  }
                }
      
                // create biblio index
                if ($sysconf['index']['type'] == 'index') {
                  $indexer->makeIndex($biblio_id ?? 0);
                }
    
                $row++;
                usleep(2500);
            }
        }
    
        utility::jsToastr('Selesai', 'Berhasil mengimpor data', 'success');
    } catch (Exception $e) {
        utility::jsToastr('Galat', $e->getMessage(), 'error');
    }

    $url = pluginUrl(reset: true);
    exit(<<<HTML
    <script>parent.$('#mainContent').simbioAJAX('{$url}')</script>
    HTML);
}

?>
<div class="menuBox">
<div class="menuBoxInner importIcon">
	<div class="per_title">
    <h2><?php echo __('Import Tool'); ?></h2>
    </div>
    <div class="infoBox">
    Impor data bibliografi dalam format Excel.
      &nbsp;<a href="<?= pluginUrl(['action' => 'download_sample']) ?>" class="s-btn btn btn-secondary notAJAX"><?= __('Download Sample') ?></a>
	  </div>
</div>
</div>
<div id="importInfo" class="infoBox" style="display: none;">&nbsp;</div><div id="importError" class="errorBox" style="display: none;">&nbsp;</div>
<?php
// create new instance
$form = new simbio_form_table_AJAX('mainForm', pluginUrl(), 'post');
$form->submit_button_attr = 'name="doImport" value="'.__('Process').'" class="btn btn-default"';

// form table attributes
$form->table_attr = 'id="dataList" class="s-table table"';
$form->table_header_attr = 'class="alterCell font-weight-bold"';
$form->table_content_attr = 'class="alterCell2"';

// csv files
$str_input  = '<div class="container-fluid">';
$str_input .= '<div class="row">';
$str_input .= '<div class="custom-file col-6">';
$str_input .= simbio_form_element::textField('file', 'importFile','','class="custom-file-input"');
$str_input .= '<label class="custom-file-label" for="customFile">Choose file</label>';
$str_input .= '</div>';
$str_input .= '<div class="col">';
$str_input .= '<div class="mt-2">Maximum '.$sysconf['max_upload'].' KB</div>';
$str_input .= '</div>';
$str_input .= '</div>';
$str_input .= '</div>';
$form->addAnything(__('File To Import'), $str_input);

// output the form
echo $form->printOut();
?>
<script>
    $('#mainForm').submit(function() {
        window.toastr.info('Moihon untuk menunggu', 'Info')
    })
</script>
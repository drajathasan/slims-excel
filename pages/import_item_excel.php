<?php
/**
 * @author Drajat Hasan
 * @email drajathasan20@gmail.com
 * @create date 2024-09-24 16:49:08
 * @modify date 2024-09-24 16:49:08
 * @desc 
 * - License : GPL-v3
 */

use SLiMS\DB;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx as Reader;
use PhpOffice\PhpSpreadsheet\Shared\Date;

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

// privileges checking
$can_read = utility::havePrivilege('bibliography', 'r');
$can_write = utility::havePrivilege('bibliography', 'w');

if (!$can_read) {
    die('<div class="errorBox">'.__('You are not authorized to view this section').'</div>');
  }

  /**
   * Download an example
   */
if (isset($_GET['action']) && $_GET['action'] == 'download_sample') {
    $spreadsheet = new Spreadsheet();
    $spreadsheet->getActiveSheet()
      ->fromArray(
          [[
            'item_code', 'call_number', 'coll_type_name',
            'inventory_code', 'received_date', 'supplier_name',
            'order_no','location_name','order_date','item_status_name',
            'site','source','invoice','price','price_currency','invoice_date',
            'input_date','last_update','title'
          ]],  // The data to set
          NULL,        // Array values with this value will not be set
          'A1'         // Top left coordinate of the worksheet range where
                       //    we want to set these values (default is A1)
      );
    $writer = new Xlsx($spreadsheet);
    $tblout ='import_item_excel_sample';
    header("Content-Type: application/xlsx");
    header("Content-Disposition: attachment; filename=$tblout.xlsx");
    header("Pragma: no-cache");
    $writer->save('php://output');
    exit;
}

/**
 * Import progress
 */
if (isset($_POST['doImport'])) {
    $pdo = DB::getInstance();
    $pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);

    try {
        $reader = new Reader;
        $reader->setReadDataOnly(true);
        $spreadsheet = $reader->load($_FILES['importFile']['tmp_name']);
        $sheet = $spreadsheet->getSheet($spreadsheet->getFirstSheetIndex());
        
        $state = $pdo->prepare(<<<SQL
        INSERT 
            IGNORE 
                INTO 
                    `item` 
                SET 
                    `biblio_id` = ?, 
                    `item_code` = ?, 
                    `call_number` = ?, 
                    `coll_type_id` = ?,
                    `inventory_code` = ?, 
                    `received_date` = ?, 
                    `supplier_id` = ?,
                    `order_no` = ?, 
                    `location_id` = ?, 
                    `order_date` = ?, 
                    `item_status_id` = ?, 
                    `site` = ?,
                    `source` = ?, 
                    `invoice` = ?, 
                    `price` = ?, 
                    `price_currency` = ?, 
                    `invoice_date` = ?,
                    `input_date` = ?, 
                    `last_update` = ?
        SQL);

        foreach ($sheet->toArray() as $data) {
            foreach ($data as $index => $field) {
                switch ($index) {
                    case 2:
                        $data[$index] = (integer)utility::getID($dbs, 'mst_coll_type', 'coll_type_id', 'coll_type_name', $data[2]);
                        break;
                    
                    case 5:
                        $data[$index] = (integer)utility::getID($dbs, 'mst_supplier', 'supplier_id', 'supplier_name', $data[5]);
                        break;

                    case 7:
                        $data[$index] = utility::getID($dbs, 'mst_location', 'location_id', 'location_name', $data[7]);
                        break;

                    case 9:
                        $data[$index] = utility::getID($dbs, 'mst_item_status', 'item_status_id', 'item_status_name', $data[9]);
                        break;

                    case 16:
                    case 17:
                        if (empty($data[$index])) $data[$index] = date('Y-m-d H:i:s');
                        break;

                    default:
                        if (empty($data[$index])) $data[$index] = null;
                        break;
                }
            }

            $title = $data[18];
            unset($data[18]);

            // get biblio_id
            $biblio = $pdo->prepare("select biblio_id from biblio where title = ?");
            $biblio->execute([$title]);

            if($biblio->rowCount() < 1) continue;
            $biblioData = $biblio->fetch(PDO::FETCH_NUM);
            $biblioId = $biblioData[0];

            $state->execute(
                array_merge(
                    [$biblioId],
                    array_values($data)
                )
            );

            if ($state->rowCount() < 1) {
                $item_code = $data[0];
                $data[4] = is_numeric($data[4]) ? date('Y-m-d', Date::excelToTimestamp($data[4])) : $data[4];
                $data[8] = is_numeric($data[8]) ? date('Y-m-d', Date::excelToTimestamp($data[8])) : $data[8];
                $data[15] = is_numeric($data[15]) ?  date('Y-m-d', Date::excelToTimestamp($data[15])) : $data[15];
                $data[16] = is_numeric($data[16]) ?  date('Y-m-d', Date::excelToTimestamp($data[16])) : $data[16];
                $data[17] = is_numeric($data[17]) ?  date('Y-m-d', Date::excelToTimestamp($data[17])) : $data[17];
                unset($data[0]);
                unset($data[16]);

                $update = $pdo->prepare(<<<SQL
                    update 
                        `item`
                    SET 
                        `call_number` = ?, 
                        `coll_type_id` = ?,
                        `inventory_code` = ?, 
                        `received_date` = ?, 
                        `supplier_id` = ?,
                        `order_no` = ?, 
                        `location_id` = ?, 
                        `order_date` = ?, 
                        `item_status_id` = ?, 
                        `site` = ?,
                        `source` = ?, 
                        `invoice` = ?, 
                        `price` = ?, 
                        `price_currency` = ?, 
                        `invoice_date` = ?,
                        `last_update` = ?
                    where
                        `item_code` = ?
                SQL);

                $update->execute(
                    array_merge(array_values($data), [$item_code])
                );
            }
        }

        utility::jsToastr('Selesai', 'Berhasil mengimpor data', 'success');
    } catch (Exception $e) {
        utility::jsToastr('Galat', $e->getMessage() . ' on line ' . $e->getLine(), 'error');
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
            <h2><?php echo __('Item Import tool'); ?></h2>
        </div>
        <div class="infoBox">
            <?php echo __('Import for item data from CSV file'); ?>
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
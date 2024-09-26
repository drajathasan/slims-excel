<?php
/**
 * @author Drajat Hasan
 * @email drajathasan20@gmail.com
 * @create date 2024-09-24 16:49:08
 * @modify date 2024-09-26 21:47:56
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

// privileges checking
$can_read = utility::havePrivilege('bibliography', 'r');
$can_write = utility::havePrivilege('bibliography', 'w');

if (!$can_read) {
  die('<div class="errorBox">'.__('You are not authorized to view this section').'</div>');
}

if (isset($_POST['doExport'])) {

    // create local function to fetch values
    function getValues($obj_db, $str_query)
    {
      // make query from database
      $_value_q = $obj_db->query($str_query);
      if ($_value_q->num_rows > 0) {
          $_value_buffer = '';
          while ($_value_d = $_value_q->fetch_row()) {
              if ($_value_d[0]) {
                  $_value_buffer .= '<'.$_value_d[0].'>';
              }
          }
          return $_value_buffer;
      }
      return null;
    }

    // pdo
    $pdo = DB::getInstance();
    $pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);

    // limit
    $limit = intval($_POST['recordNum']);
    $offset = intval($_POST['recordOffset']);

    // fetch all data from biblio table
    $sql = "SELECT
            i.item_code, i.call_number, ct.coll_type_name,
            i.inventory_code, i.received_date, spl.supplier_name,
            i.order_no, loc.location_name,
            i.order_date, st.item_status_name, i.site,
            i.source, i.invoice, i.price, i.price_currency, i.invoice_date,
            i.input_date, i.last_update, b.title
            FROM item AS i
            LEFT JOIN biblio AS b ON i.biblio_id=b.biblio_id
            LEFT JOIN mst_coll_type AS ct ON i.coll_type_id=ct.coll_type_id
            LEFT JOIN mst_supplier AS spl ON i.supplier_id=spl.supplier_id
            LEFT JOIN mst_item_status AS st ON i.item_status_id=st.item_status_id
            LEFT JOIN mst_location AS loc ON i.location_id=loc.location_id";
    if ($limit > 0) { $sql .= ' LIMIT '.$limit; }
    if ($offset > 1) {
      if ($limit > 0) {
          $sql .= ' OFFSET '.($offset-1);
      } else {
          $sql .= ' LIMIT '.($offset-1).',99999999999';
      }
    }

    try {
        $items = DB::getInstance()->query($sql);

        if ($items->rowCount() < 1) throw new Exception("Data eksemplar masih kosong!");

        $row = 0;
        $headers = [];
        $body = [];
        while ($item = $items->fetch(PDO::FETCH_ASSOC)) {
            
            if ($row === 0) {
                $headers = array_keys($item);
            }

            $body[] = array_values($item);
            $row++;
        }

        $data = array_merge([$headers], $body);

        $spreadsheet = new Spreadsheet();
        $spreadsheet->getActiveSheet()
        ->fromArray(
            $data,  // The data to set
            NULL,        // Array values with this value will not be set
            'A1'         // Top left coordinate of the worksheet range where
                        //    we want to set these values (default is A1)
        );
        $writer = new Xlsx($spreadsheet);
        $tblout ='export_item_excel';
        header("Content-Type: application/xlsx");
        header("Content-Disposition: attachment; filename=$tblout.xlsx");
        header("Pragma: no-cache");
        $writer->save('php://output');
    } catch (Exception $e) {
        utility::jsToastr('Galat', $e->getMessage() . ' - ' . $e->getLine(), 'error');
    }
    exit;
}
?>
<div class="menuBox">
<div class="menuBoxInner exportIcon">
	<div class="per_title">
    	<h2><?php echo __('Export Tool'); ?></h2>
	</div>
	<div class="infoBox">
    	Ekspor data bibitemliografi dalam format Excel
	</div>
</div>
</div>
<?php
// create new instance
$form = new simbio_form_table_AJAX('mainForm', pluginUrl(), 'post');
$form->submit_button_attr = 'name="doExport" value="'.__('Process').'" class="btn btn-default"';

// form table attributes
$form->table_attr = 'id="dataList" class="s-table table"';
$form->table_header_attr = 'class="alterCell font-weight-bold"';
$form->table_content_attr = 'class="alterCell2"';

// csv files
$form->addTextField('text', 'recordNum', __('Number of Records To Export (0 for all records)'), '0', 'style="width: 10%;" class="form-control"');
$form->addTextField('text', 'recordOffset', __('Start From Record'), '1', 'style="width: 10%;"  class="form-control"');

// output the form
echo $form->printOut();
?>
<script>
    $('#mainForm').submit(function() {
        window.toastr.info('Moihon untuk menunggu', 'Info')
        setTimeout(() => {
            $('.loader').attr('style', 'display: none !important');
        }, 2500);
    })
</script>
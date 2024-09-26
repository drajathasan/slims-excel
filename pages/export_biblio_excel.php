<?php
/**
 * @author Drajat Hasan
 * @email drajathasan20@gmail.com
 * @create date 2024-09-24 16:49:08
 * @modify date 2024-09-26 21:41:40
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
        b.biblio_id, b.title, gmd.gmd_name, b.edition,
        b.isbn_issn, publ.publisher_name, b.publish_year,
        b.collation, b.series_title, b.call_number,
        lang.language_name, pl.place_name, b.classification,
        b.notes, b.image, b.sor
        FROM biblio AS b
        LEFT JOIN mst_gmd AS gmd ON b.gmd_id=gmd.gmd_id
        LEFT JOIN mst_publisher AS publ ON b.publisher_id=publ.publisher_id
        LEFT JOIN mst_language AS lang ON b.language_id=lang.language_id
        LEFT JOIN mst_place AS pl ON b.publish_place_id=pl.place_id ORDER BY b.last_update DESC";
    if ($limit > 0) { $sql .= ' LIMIT '.$limit; }
    if ($offset > 1) {
      if ($limit > 0) {
          $sql .= ' OFFSET '.($offset-1);
      } else {
          $sql .= ' LIMIT '.($offset-1).',99999999999';
      }
    }

    try {
        $biblios = DB::getInstance()->query($sql);

        if ($biblios->rowCount() < 1) throw new Exception("Data biblio masih kosong!");

        $row = 0;
        $headers = [];
        $body = [];
        while ($biblio = $biblios->fetch(PDO::FETCH_ASSOC)) {
            $id = $biblio['biblio_id'];
            unset($biblio['biblio_id']);
            
            if ($row === 0) {
                $headers = array_keys($biblio);
                $headers[] = 'authors';
                $headers[] = 'topics';
                $headers[] = 'item_code';
            }

            $biblioValues = array_values($biblio);

            // Author
            $biblioValues[] = getValues($dbs, 'SELECT a.author_name FROM biblio_author AS ba
              LEFT JOIN mst_author AS a ON ba.author_id=a.author_id
              WHERE ba.biblio_id='.$id)??'';

            // topics
            $biblioValues[] = getValues($dbs, 'SELECT t.topic FROM biblio_topic AS bt 
            LEFT JOIN mst_topic AS t ON bt.topic_id=t.topic_id 
            WHERE bt.biblio_id='.$id)??'';

            // item code
            $biblioValues[] = getValues($dbs, 'SELECT item_code FROM item AS i
              WHERE i.biblio_id='.$id)??'';

            $body[] = $biblioValues;
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
        $tblout ='export_biblio_excel';
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
    	Ekspor data bibliografi dalam format Excel
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
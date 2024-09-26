<?php
use SLiMS\DB;

$backup_log = DB::getInstance()->query('select backup_log_id from backup_log where date(backup_time) = \'' . (date('Y-m-d')) . '\'');

if ($backup_log->rowCount() < 1) {
    $can_read = utility::havePrivilege('system', 'r');
    $can_write = utility::havePrivilege('system', 'w');
?>
    <div class="d-flex flex-column align-items-center" style="height: 100vh">
        <img style="width: 450px" src="data:image/png;base64, <?= base64_encode(file_get_contents(__DIR__ . '/../static/images/Warning-rafiki.png')) ?>"/>
        <h3 class="font-weight-bold">Yah!</h3>
        <?php
        if (!($can_read AND $can_write)) {
            ?>
            <p style="width: 450px; text-align: center; font-size: 12pt">Hari ini SLiMS kamu <em>database</em> nya belum di <em>backup</em> segera hubungi Admin untuk melakukan nya.</p>
            <?php
        } else {
        ?>
            <p style="width: 450px; text-align: center; font-size: 12pt">Kamu belum melakukan <em>backup database</em> hari ini. Lakukan terlebih dahulu agar jika terjadi sesuatu anomali maka kamu masih punya data yang masih bagus.</p>
            <div class="d-flex flex-row" style="gap: 5px;">
                <a href="<?= pluginUrl(['action' => 'bypass_backup']) ?>" class="btn btn-outline-secondary">Gak perlu</a>
                <a href="<?= MWB ?>system/backup.php" class="btn btn-primary">Backup Sekarang</a>
            </div>
        <?php
        }
        ?>
    </div>
<?php
    exit;
}
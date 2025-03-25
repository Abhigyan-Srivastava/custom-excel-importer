<?php
/**
 * Plugin Name: Excel Post Importer
 * Description: Import post data using an Excel sheet.
 * Version: 1.0.0
 * Author: Abhigyan
 */

if (!defined('ABSPATH')) {
    exit;
}

require_once __DIR__ . '/vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\IOFactory;

function epi_admin_menu() {
    add_menu_page('Excel Post Importer', 'Post Importer', 'manage_options', 'excel-post-importer', 'epi_import_page');
}
add_action('admin_menu', 'epi_admin_menu');

function epi_import_page() {
    ?>
    <div class="wrap">
        <h1>Import Posts from Excel</h1>
        <form method="post" enctype="multipart/form-data">
            <input type="file" name="excel_file" required>
            <input type="hidden" name="epi_import_nonce" value="<?php echo wp_create_nonce('epi_import_action'); ?>">
            <button type="submit" name="import" class="button button-primary">Import</button>
        </form>
    </div>
    <?php
    if (isset($_POST['import']) && check_admin_referer('epi_import_action', 'epi_import_nonce')) {
        epi_process_import();
    }
}

function epi_process_import() {
    if (!isset($_FILES['excel_file']) || $_FILES['excel_file']['error'] != UPLOAD_ERR_OK) {
        echo '<div class="error"><p>Error uploading file.</p></div>';
        return;
    }

    $file = $_FILES['excel_file']['tmp_name'];
    $spreadsheet = IOFactory::load($file);
    $worksheet = $spreadsheet->getActiveSheet();
    $rows = $worksheet->toArray();

    echo($rows);

    foreach ($rows as $index => $row) {
        if ($index == 0) continue;

        echo($row[0]);
        echo($row[1]);

        $post_data = [
            'post_title'   => sanitize_text_field($row[0]),
            'post_content' => sanitize_textarea_field($row[1]),
            'post_status'  => 'publish',
            'post_type'    => 'post',
        ];

        $existing_post = get_page_by_title($post_data['post_title'], OBJECT, 'post');
        if ($existing_post) {
            $post_data['ID'] = $existing_post->ID;
        }

        wp_insert_post($post_data);
    }
    echo '<div class="updated"><p>Posts imported successfully!</p></div>';
}

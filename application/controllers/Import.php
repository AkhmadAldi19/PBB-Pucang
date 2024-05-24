<?php
defined('BASEPATH') or exit('No direct script access allowed');

//load Spout Library
require_once APPPATH . 'third_party/spout/src/Spout/Autoloader/autoload.php';

use Box\Spout\Reader\Common\Creator\ReaderEntityFactory;

class Import extends CI_Controller
{

    function __construct()
    {
        parent::__construct();
        //load model
        $this->load->model('app');
    }

    public function index()
    {
        //ketika button submit diklik
        if ($this->input->post('submit', TRUE) == 'upload') {
            $config['upload_path']      = './temp_doc/'; //siapkan path untuk upload file
            $config['allowed_types']    = 'xlsx|xls'; //siapkan format file
            $config['file_name']        = 'doc' . time(); //rename file yang diupload
            

            $this->load->library('upload', $config);

            if ($this->upload->do_upload('excel')) {
                //fetch data upload
                $file   = $this->upload->data();

                $reader = ReaderEntityFactory::createXLSXReader(); //buat xlsx reader
                $reader->open('temp_doc/' . $file['file_name']); //open file xlsx yang baru saja diunggah

                //looping pembacaat sheet dalam file        
                foreach ($reader->getSheetIterator() as $sheet) {
                    $numRow = 1;

                    //siapkan variabel array kosong untuk menampung variabel array data
                    $save   = array();

                    //looping pembacaan row dalam sheet
                    foreach ($sheet->getRowIterator() as $row) {

                        if ($numRow > 1) {
                            //ambil cell
                            $cells = $row->getCells();

                            $data = array(
                                'id'                => $cells[0],
                                'nop'               => $cells[1],
                                'nama'              => $cells[2],
                                'bumi'              => $cells[3],
                                'bangunan'          => $cells[4],
                                'pajak'             => $cells[5],
                                'alamat_op'         => $cells[6],
                                'alamat_wp'         => $cells[7],
                                'ket'               => $cells[8]
                            );

                            //tambahkan array $data ke $save
                            array_push($save, $data);
                        }

                        $numRow++;
                    }
                    //simpan data ke database
                    $this->app->simpan($save);

                    //tutup spout reader
                    $reader->close();

                    //hapus file yang sudah diupload
                    unlink('temp_doc/' . $file['file_name']);

                    //tampilkan pesan success dan redirect ulang ke index controller import
                    echo    '<script type="text/javascript">
                               alert(\'Data berhasil disimpan\');
                               window.location.replace("' . base_url() . '");
                           </script>';
                }
            } else {
                echo "Error :" . $this->upload->display_errors(); //tampilkan pesan error jika file gagal diupload
            }
        }
        
        $this->load->view('import');
    }
}

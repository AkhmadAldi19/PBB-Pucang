<?php
defined('BASEPATH') OR exit('No direct script access allowed');
require 'vendor/autoload.php';

class Welcome extends CI_Controller {

    public function index()
    {
        if($_SERVER['REQUEST_METHOD']=='POST')
        {
            $upload_status =  $this->uploadDoc();
            if($upload_status != false)
            {
                $inputFileName = 'assets/uploads/imports/'.$upload_status;
                $inputFileType = \PhpOffice\PhpSpreadsheet\IOFactory::identify($inputFileName);
                $reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReader($inputFileType);
                $spreadsheet = $reader->load($inputFileName);
                $sheet = $spreadsheet->getSheet(0);
                $count_Rows = 0;
                foreach($sheet->getRowIterator() as $row)
                {
                    // Mendapatkan nilai sel-sel pada setiap baris
                    $id = $spreadsheet->getActiveSheet()->getCell('A'.$row->getRowIndex());
                    $nop = $spreadsheet->getActiveSheet()->getCell('B'.$row->getRowIndex());
                    $nama = $spreadsheet->getActiveSheet()->getCell('C'.$row->getRowIndex());
                    $bumi = $spreadsheet->getActiveSheet()->getCell('D'.$row->getRowIndex());
                    $bangunan = $spreadsheet->getActiveSheet()->getCell('E'.$row->getRowIndex());
                    $pajak = $spreadsheet->getActiveSheet()->getCell('F'.$row->getRowIndex());
                    $alamat_op = $spreadsheet->getActiveSheet()->getCell('G'.$row->getRowIndex());
                    $alamat_wp = $spreadsheet->getActiveSheet()->getCell('H'.$row->getRowIndex());
                    $ket = $spreadsheet->getActiveSheet()->getCell('I'.$row->getRowIndex());
                    $nomor_hp = $spreadsheet->getActiveSheet()->getCell('J'.$row->getRowIndex());
                    $nama_petugas = $spreadsheet->getActiveSheet()->getCell('K'.$row->getRowIndex());
                    
                    // Menyimpan data ke dalam array
                    $data = array(
                        'id' => $id,
                        'nop' => $nop,
                        'nama' => $nama,
                        'bumi' => $bumi,
                        'bangunan' => $bangunan,
                        'pajak' => $pajak,
                        'alamat_op' => $alamat_op,
                        'alamat_wp' => $alamat_wp,
                        'ket' => $ket,
                        'nomor_hp' => $nomor_hp,
                        'nama_petugas' => $nama_petugas
                    );

                    // Menyimpan data ke dalam tabel data_pbb
                    $this->db->insert('data_pbb', $data);
                    $count_Rows++;
                }
                $this->session->set_flashdata('success', 'Successfully Data Imported');
                redirect(base_url());
            }
            else
            {
                $this->session->set_flashdata('error', 'File is not uploaded');
                redirect(base_url());
            }
        }
        else
        {
            $this->load->view('import');
        }
        
    }

    function uploadDoc()
    {
        $uploadPath = 'assets/uploads/imports/';
        if(!is_dir($uploadPath))
        {
            mkdir($uploadPath, 0777, TRUE); // UNTUK MEMBUAT DIRECTORY JIKA TIDAK ADA
        }

        $config['upload_path'] = $uploadPath;
        $config['allowed_types'] = 'csv|xlsx|xls';
        $config['max_size'] = 1000000;
        $this->load->library('upload', $config);
        $this->upload->initialize($config);
        if($this->upload->do_upload('upload_excel'))
        {
            $fileData = $this->upload->data();
            return $fileData['file_name'];
        }
        else
        {
            return false;
        }
    }

}
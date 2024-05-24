<?php
defined('BASEPATH') OR exit('No direct script access allowed');

class Upload extends CI_Controller {

    public function __construct(){

        parent::__construct();

        // load base_url
        $this->load->helper('url');

        // Load Model
        $this->load->model('Import_model');
    }
    
    public function index(){

        // Check form submit or not 
        if($this->input->post('upload') != NULL ){ 
            $data = array(); 
            if(!empty($_FILES['file']['name'])){ 
                // Set preference 
                $config['upload_path'] = 'assets/files/'; 
                $config['allowed_types'] = 'csv'; 
                $config['max_size'] = '1000'; // max_size in kb 
                $config['file_name'] = $_FILES['file']['name']; 

                // Load upload library 
                $this->load->library('upload',$config); 
   
                // File upload
                if($this->upload->do_upload('file')){ 
                    // Get data about the file
                    $uploadData = $this->upload->data(); 
                    $filename = $uploadData['file_name']; 

                    // Reading file
                    $file = fopen("assets/files/".$filename,"r");
                    $i = 0;
                    $numberOfFields = 8; // Total number of fields
                    $importData_arr = array();
                       
                    while (($filedata = fgetcsv($file, 10000, ",")) !== FALSE) {
                        $num = count($filedata);

                        if($numberOfFields == $num){
                            for ($c=0; $c < $num; $c++) {
                                $importData_arr[$i][] = $filedata [$c];
                            }
                        }
                        $i++;
                    }
                    fclose($file);

                    $skip = 0;

                    // insert import data
                    foreach($importData_arr as $datadhkp){
                        
                        // Skip first row
                        if($skip != 0){
                            $this->Import_model->insertRecord($datadhkp);
                        }
                        $skip ++;
                    }
                    $data['response'] = 'successfully uploaded '.$filename; 
                }else{ 
                    $data['response'] = 'failed'; 
                } 
            }else{ 
                $data['response'] = 'failed'; 
            } 
            // load view
            $data['content'] = 'contents/upload/index';
            $this->load->view('app', $data); 
            
        }else{
            // load view 
            $data['content'] = 'contents/upload/index';
            $this->load->view('app', $data); 
        } 

    }

    
}

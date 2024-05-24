<?php if ( ! defined('BASEPATH')) exit('No direct script access allowed');

class Import_model extends CI_Model {

    function insertRecord($record){
        
        if(count($record) > 0){
            
            // Check user
            $this->db->select('*');
            $this->db->where('nop', $record[0]);
            $q = $this->db->get('nama');
            $response = $q->result_array();
            
            // Insert record
            if(count($response) == 0){
                $datadhkp = array(
                    "nop" => trim($record[0]),
                    "nama" => trim($record[1]),
                    "bumi" => trim($record[2]),
                    "bangunan" => trim($record[3]),
                    "pajak" => trim($record[4]),
                    "alamat_op" => trim($record[5]),
                    "alamat_wp" => trim($record[6]),
                    "ket" => trim($record[7])
                );

                $this->db->insert('data_pbb', $datadhkp);
            }
            
        }
        
    }

}
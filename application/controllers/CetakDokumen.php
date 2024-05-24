<?php
defined('BASEPATH') OR exit('No direct script access allowed');

class CetakDokumen extends CI_Controller {

    private $CI;

    public function __construct(){

		parent :: __construct();
        if($this->session->userdata('username') == ''){
            redirect(base_url() . 'login');
		}
        $this->load->model("cetakdokumenmodel");
        $this->load->library('Pdf');
    }

    public function index(){
        $data['datas'] = $this->cetakdokumenmodel->fetchAll();
        
        $data['rupiah'] = function($angka){
	
            $hasil_rupiah = "Rp " . number_format($angka,0,',','.');
            return $hasil_rupiah;
         
        };
        

        $this->load->view('contents/cetak/pbbterhutang', $data);
    }
 
    public function pbblunas(){
        $data['datas'] = $this->cetakdokumenmodel->fetchAllLunas($this->session->userdata('id'));


        $data['rupiah'] = function($angka){
	
            $hasil_rupiah = "Rp " . number_format($angka,0,',','.');
            return $hasil_rupiah;
         
        };

        $this->load->view('contents/cetak/pbblunas', $data);
    }

    public function pbbsetor(){
        if(!empty($this->uri->segment(3)) && !empty($this->uri->segment(4))){
            $data['tgl_awal'] = $this->uri->segment(3);
            $data['tgl_akhir'] = $this->uri->segment(4);
            $data['datas'] = $this->cetakdokumenmodel->fetchSetor($this->session->userdata('id'), $data['tgl_awal'], $data['tgl_akhir']);

            $data['rupiah'] = function($angka){
        
                $hasil_rupiah = "Rp " . number_format($angka,0,',','.');
                return $hasil_rupiah;
             
            };
    
            $this->load->view('contents/cetak/pbbsetor', $data);

        }
    }

    public function tandaTerima(){
        if(!empty($this->uri->segment(3))){

            $id = $this->uri->segment(3);
            $data['data'] = $this->cetakdokumenmodel->fetchPBB($id);
            $data['rupiah'] = function($angka){
        
                $hasil_rupiah = "Rp " . number_format($angka,0,',','.');
                return $hasil_rupiah;
             
            };

            $data['tanggal'] = function($tanggal){
                $arrtgl = explode("-", $tanggal);
                $arrbulan = array(
                    'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'
                );
                return $arrtgl[2] . ' ' . $arrbulan[$arrtgl[1]-1] . ' ' . $arrtgl[0];
            };

            $this->load->view('contents/cetak/tandaterima', $data);
        }

    }
}
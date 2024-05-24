<?php
defined('BASEPATH') OR exit('No direct script access allowed');

class RekapitulasiPBB extends CI_Controller {

    public function __construct(){

		parent :: __construct();
        if($this->session->userdata('username') == ''){
            redirect(base_url() . 'login');
		}
        $this->load->model("rekapmodel");
	}

	public function index(){

        $data['content'] = 'contents/rekapitulasipbb/index';
        $data['data'] = $this->rekapmodel->fetchData();
        $data['rupiah'] = function($angka){
            $hasil_rupiah = "Rp " . number_format($angka,0,',','.');
            return $hasil_rupiah;
        };
        $data['title'] = 'Rekapitulasi PBB';
		$this->load->view('app', $data);
    }
    public function rupiah($angka){
	
        $hasil_rupiah = "Rp " . number_format($angka,0,',','.');
        return $hasil_rupiah;
     
    }
}
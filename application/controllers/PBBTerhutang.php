<?php
defined('BASEPATH') OR exit('No direct script access allowed');

class PBBTerhutang extends CI_Controller {


    public function __construct(){

		parent :: __construct();
        if($this->session->userdata('username') == ''){
            redirect(base_url() . 'login');
		}
        $this->load->model("pbbterhutangmodel");
	}

	public function index(){

        //pagination
        $config['base_url'] = site_url('datasppt/index');
		$config['total_rows'] = $this->db->count_all('data_pbb');
		$config['per_page'] = 5000;
        $config['uri_segment'] = 3;
		$choice = $config['total_rows'] / $config['per_page'];
		$config['num_links'] = 3;

		$config['first_link']       = 'First';
        $config['last_link']        = 'Last';
        $config['next_link']        = 'Next';
        $config['prev_link']        = 'Prev';
        $config['full_tag_open']    = '<div class="pagging text-center"><nav><ul class="pagination justify-content-center">';
        $config['full_tag_close']   = '</ul></nav></div>';
        $config['num_tag_open']     = '<li class="page-item"><span class="page-link">';
        $config['num_tag_close']    = '</span></li>';
        $config['cur_tag_open']     = '<li class="page-item active"><span class="page-link">';
        $config['cur_tag_close']    = '<span class="sr-only">(current)</span></span></li>';
        $config['next_tag_open']    = '<li class="page-item"><span class="page-link">';
        $config['next_tagl_close']  = '<span aria-hidden="true">&raquo;</span></span></li>';
        $config['prev_tag_open']    = '<li class="page-item"><span class="page-link">';
        $config['prev_tagl_close']  = '</span>Next</li>';
        $config['first_tag_open']   = '<li class="page-item"><span class="page-link">';
        $config['first_tagl_close'] = '</span></li>';
        $config['last_tag_open']    = '<li class="page-item"><span class="page-link">';
		$config['last_tagl_close']  = '</span></li>';
		
		$this->pagination->initialize($config);
		$data['page'] = ($this->uri->segment(3)) ? $this->uri->segment(3) : 0;

		$data['datas'] = $this->pbbterhutangmodel->fetchData($config['per_page'], $data['page']);

        $data['pagination'] = $this->pagination->create_links();

        $data['content'] = 'contents/pbbterhutang/index';
        $data['title'] = 'PBB Belum Bayar';

		$this->load->view('app', $data);
    }
    
    public function bayarPajak(){
        $id = $this->uri->segment(3);

        $data['data'] = $this->pbbterhutangmodel->fetchSingle($id);
        $data['content'] = 'contents/pbbterhutang/pajak';

		$this->load->view('app', $data);
    }

    public function bayarPajakSave(){
        
        $this->form_validation->set_rules('nop', 'NOP', 'required|numeric');
        $this->form_validation->set_rules('nama', 'Nama', 'required');
        $this->form_validation->set_rules('pajak', 'Pajak', 'required|numeric');
        if ($this->form_validation->run() == TRUE)
        {
            $data = array(
                'nop'=>$this->input->post('nop'),
                'nama'=>$this->input->post('nama'),
                'pajak'=>$this->input->post('pajak'),
                'denda'=>$this->input->post('denda'),
                'total_pajak'=>$this->input->post('pajak') + $this->input->post('denda'),
                'tanggal'=>date('Y-m-d'),
                'ket'=>'Lunas',
                'id_petugas'=>$this->session->userdata('id'),
                'nm_petugas'=>$this->session->userdata('username'),
                'id_penyetor'=>0
            );

            //update datapbb
            $datapbb= array(
                'id'=>$this->uri->segment(3),
                'ket'=>'Lunas'
            );

            $this->pbbterhutangmodel->insertDataBayar($data);
            $this->pbbterhutangmodel->updateData($datapbb);

            redirect(base_url() . 'pbbterhutang/bayar');
        }else{
            $this->bayarPajak();
        }
    }

    public function bayar(){
        $this->index();
    }
}
<?php
defined('BASEPATH') OR exit('No direct script access allowed');

class SetorPBB extends CI_Controller {


    public function __construct(){

		parent :: __construct();
        if($this->session->userdata('username') == ''){
            redirect(base_url() . 'login');
		}
        $this->load->model("setorpbbmodel");
    }

    
	public function index(){

        //pagination
        $config['base_url'] = site_url('setorpbb/index');
		$config['total_rows'] = $this->db->where('id_admin',$this->session->userdata('id'))->from('data_setor')->count_all_results();
		$config['per_page'] = 10;
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

		$data['datas'] = $this->setorpbbmodel->fetchData($config['per_page'], $data['page'], $this->session->userdata('id'));

        $data['pagination'] = $this->pagination->create_links();

        $data['content'] = 'contents/setorpbb/index';
        $data['script'] = 'contents/setorpbb/script';
        $data['title'] = 'Setor PBB';

		$this->load->view('app', $data);
    }

}
<?php
defined('BASEPATH') OR exit('No direct script access allowed');

class Petugas extends CI_Controller {


    public function __construct(){

		parent :: __construct();
        if($this->session->userdata('username') == ''){
            redirect(base_url() . 'login');
		}
        $this->load->model("petugasmodel");
    }
    
    public function index(){
        
        //pagination
        $config['base_url'] = site_url('petugas/index');
		$config['total_rows'] = $this->db->count_all('admin');
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

		$data['datas'] = $this->petugasmodel->fetchData($config['per_page'], $data['page']);

        $data['pagination'] = $this->pagination->create_links();

        $data['content'] = 'contents/petugas/index';

		$this->load->view('app', $data);
    }

        public function addForm(){
        
        $data['content'] = 'contents/petugas/add';
        $data['title'] = 'Tambah Data Pengguna';
        $this->load->view('app', $data);
    }
    

    public function addSave(){
        $config['file_name'] = '1.jpg';
        $config['upload_path'] = './uploads/myassets/';
        $config['allowed_types'] = 'gif|jpg|jpeg|png';
        $config['max_size'] = 10000;

        $this->load->library('upload', $config);

        $this->upload->overwrite = true;
        $this->form_validation->set_rules('uname', 'User Name', 'required');
        $this->form_validation->set_rules('hp', 'Nomor HP', 'required|numeric');
        $this->form_validation->set_rules('pass', 'Kata Sandi', 'required');
        $this->form_validation->set_rules('level', 'Level Petugas', 'required');
        if ($this->form_validation->run() == TRUE)
        {
            $data = array(
                'uname'=>$this->input->post('uname'),
                'hp'=>$this->input->post('hp'),
                'pass'=>md5($this->input->post('pass')),
                'level'=>$this->input->post('level'),
                'foto'=>$this->input->post('foto'),
                'tgl'=>date('Y-m-d')
                
                
            );
                // print_r($data);
                $this->petugasmodel->insertData($data);
                redirect(base_url() . 'petugas/added');
        }else{
            $this->addForm();
        }
    }

    public function added(){
        $this->index();
    }

       public function editForm(){
        $id = $this->uri->segment(3);

        $data['data'] = $this->petugasmodel->fetchSingle($id);
        $data['content'] = 'contents/petugas/edit';
        $data['title'] = 'Edit Wajib Pajak';
        $this->load->view('app', $data);

    }
    public function deleteData(){
        $id = $this->uri->segment(3);
        
        $this->petugasmodel->deleteData($id);
        redirect(base_url() . 'petugas/deleted');

    }

    public function deleted(){
        $this->index();
    }
    
    public function editSave(){
        $config['file_name'] = '1.jpg';
        $config['upload_path'] = './uploads/myassets/';
        $config['allowed_types'] = 'gif|jpg|jpeg|png';
        $config['max_size'] = 10000;

        $this->load->library('upload', $config);
        $this->upload->overwrite = true;
        

        $this->form_validation->set_rules('uname', 'User Name', 'required');
        $this->form_validation->set_rules('hp', 'Nomo HP', 'required');
        $this->form_validation->set_rules('pass', 'Kata Sandi', 'required');
        $this->form_validation->set_rules('foto', 'Foto Diri', 'required');
        $this->form_validation->set_rules('level', 'Level Akun', 'required');

        if ($this->form_validation->run() == TRUE)
        {
            if ($this->upload->do_upload('foto')){

                $data = array(
                    'id'=>$this->uri->segment(3),
                    'uname'=>$this->input->post('uname'),
                    'hp'=>$this->input->post('hp'),
                    'pass'=>md5($this->input->post('pass')),
                    'level'=>$this->input->post('level'),
                    'foto'=>$this->input->post('foto'),
                    'tgl'=>date('Y-m-d')
                );

                // print_r($data);
                $this->profilmodel->updateData($data);
                redirect(base_url() . 'petugas/added');
            }else{
                $this->editForm();
            }
        }else{
            $this->editForm();
        }

    }
}
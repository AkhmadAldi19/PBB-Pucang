<?php
defined('BASEPATH') OR exit('No direct script access allowed');

class Profil extends CI_Controller {


    public function __construct(){

		parent :: __construct();
        if($this->session->userdata('username') == ''){
            redirect(base_url() . 'login');
		}
        $this->load->model("profilmodel");
	}

	public function index()
	{
        $data['content'] = 'contents/profil/index';
        $data['data'] = $this->profilmodel->fetchSingle();
        $data['title'] = 'Profil Desa';

		$this->load->view('app', $data);
    }
    
	public function editForm()
	{
        $data['content'] = 'contents/profil/edit';
        $data['data'] = $this->profilmodel->fetchSingle();
        $data['title'] = 'Perbarui Profil Desa';

		$this->load->view('app', $data);
    }
    
    public function editSave(){
		$config['file_name'] = 'logo';
		$config['upload_path'] = './uploads/myassets/';
		$config['allowed_types'] = 'gif|jpg|jpeg|png';
		$config['max_size'] = 10000;

		$this->load->library('upload', $config);
        $this->upload->overwrite = true;
        

        $this->form_validation->set_rules('nama_desa', 'Namas Desa', 'required');
        $this->form_validation->set_rules('kec', 'Kecamatan', 'required');
        $this->form_validation->set_rules('kab', 'Kabupaten', 'required');
        $this->form_validation->set_rules('kades', 'Kepala Desa', 'required');
        $this->form_validation->set_rules('kantor', 'Alamat Kantor', 'required');
        $this->form_validation->set_rules('email', 'Email', 'required');
        $this->form_validation->set_rules('web', 'Web', 'required');
        $this->form_validation->set_rules('rek', 'No Rekening', 'required');
        $this->form_validation->set_rules('kop1', 'Kop1', 'required');
        $this->form_validation->set_rules('kop2', 'Kop2', 'required');
        $this->form_validation->set_rules('kop3', 'Kop3', 'required');

        if ($this->form_validation->run() == TRUE)
        {
			if ($this->upload->do_upload('logo')){

                $data = array(
                    'id'=>$this->uri->segment(3),
                    'nama_desa'=>$this->input->post('nama_desa'),
                    'kec'=>$this->input->post('kec'),
                    'kab'=>$this->input->post('kab'),
                    'kades'=>$this->input->post('kades'),
                    'kantor'=>$this->input->post('kantor'),
                    'email'=>$this->input->post('email'),
                    'web'=>$this->input->post('web'),
                    'rek'=>$this->input->post('rek'),
                    'kop1'=>$this->input->post('kop1'),
                    'kop2'=>$this->input->post('kop2'),
                    'kop3'=>$this->input->post('kop3'),
                    'logo'=>$this->upload->data('file_name')
                );

                // print_r($data);
                $this->profilmodel->updateData($data);
                redirect(base_url() . 'profil/index');
            }else{
                $this->editForm();
            }
        }else{
            $this->editForm();
        }

    }
}

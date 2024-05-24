<?php
defined('BASEPATH') OR exit('No direct script access allowed');

class PBBLunas extends CI_Controller {


    public function __construct(){

		parent :: __construct();
        if($this->session->userdata('username') == ''){
            redirect(base_url() . 'login');
		}
        $this->load->model("pbblunasmodel");
    }

    public function index(){
        //pagination
        $config['base_url'] = site_url('pbblunas/index');
		$config['total_rows'] = $this->db->where('ket', 'lunas')->where('id_petugas', $this->session->userdata('id'))->from('data_bayar')->count_all_results();
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

		$data['datas'] = $this->pbblunasmodel->fetchData($config['per_page'], $data['page'],$this->session->userdata('id'));

        $data['pagination'] = $this->pagination->create_links();

        $data['content'] = 'contents/pbblunas/index';
        $data['script'] = 'contents/pbblunas/script';
        $data['title'] = 'PBB Lunas';

		$this->load->view('app', $data);
    }

    public function setor(){
        
        // $this->pbblunasmodel->updateBatch($dataBatch);
        
        if($this->input->post('type')=='setor'){
            $dataBatch = array();
            for($i = 0; $i < count($this->input->post('collection')['id']); $i++){
                $myCol = explode(';', $this->input->post('collection')['id'][$i]);
                $dataBatch[$i] = array(
                    'id'=>$myCol[0],
                    'ket'=>'Setor'
                );
            }
    
            $dataBatchSetor = array();
            for($i = 0; $i < count($this->input->post('collection')['id']); $i++){
                $myCol = explode(';', $this->input->post('collection')['id'][$i]);
                $dataBatchSetor[$i] = array(
                    'kode'=>'1801',
                    'nop'=>$myCol[1],
                    'total_pajak'=>$myCol[2],
                    'tahun'=>date('Y'),
                    'rek'=>'4020304050623',
                    'tanggal'=>date('Y-m-d'),
                    'id_admin'=>$this->session->userdata('id'),
                    'status'=>1
                );
            }

            $this->db->update_batch('data_bayar', $dataBatch, 'id');
            $this->db->insert_batch('data_setor', $dataBatchSetor);
            
            redirect(base_url() . 'pbblunas/setored');

            // print_r($this->input->post());
            // echo '<br>';
            // print_r($dataBatch);
            // echo '<br>';
            // print_r($dataBatchSetor);
        }else if($this->input->post('type')=='batal'){
            $ids = array();
            for($i = 0; $i < count($this->input->post('collection')['id']); $i++){
                $myCol = explode(';', $this->input->post('collection')['id'][$i]);
                array_push($ids, $myCol[0]);

            }
            print_r($ids);

            $this->db->where_in('id', $ids);
            $this->db->delete('data_bayar');

            redirect(base_url() . 'pbblunas/bataled');
        }
    }

    public function setored(){
        $this->index();

    }

    public function bataled(){
        $this->index();

    }
}
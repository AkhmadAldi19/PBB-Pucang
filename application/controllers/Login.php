<?php

defined('BASEPATH') or exit('No direct script access allowed');

class Login extends CI_Controller
{

    public function __construct()
    {

        parent::__construct();
        // if($this->session->userdata('username') == ''){
        //     redirect(base_url() . 'login');
        // }
        $this->load->model("loginmodel");
    }

    public function index()
    {
        $this->load->view('contents/login/index');
    }

    public function login_validation()
    {
        $data = array(
            'username' => $this->input->post('username'),
            'password' => md5($this->input->post('password')),
        );

        $user = $this->loginmodel->fetch_user($data)->row();
        if (!empty($user)) {
            $session_data = array(
                'id' => $user->id,
                'username' => $user->uname,
                'hp' => $user->hp,
                'foto' => $user->foto
            );
            $this->session->set_userdata($session_data);
            redirect(base_url() . 'dashboard');
        } else {
            redirect(base_url() . 'login');
        }
    }

    function logout()
    {
        $this->session->unset_userdata('username');
        $this->session->unset_userdata('id');
        redirect(base_url() . 'login');
    }
}

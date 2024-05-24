<?php
class LoginModel extends CI_Model{
    function fetch_user($data){
        $query = $this->db->query("SELECT * FROM admin WHERE uname = '". $data['username'] ."' AND pass= '". $data['password'] ."';");

        return $query;
    }
}
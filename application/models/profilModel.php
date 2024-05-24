<?php
class profilModel extends CI_Model{
    public function fetchSingle(){
        return $this->db->get_where('profil', array('id' => 1));
    }

    public function updateData($data){
        $this->db->where("id", $data['id']);
        $this->db->update("profil", $data);
    }
}
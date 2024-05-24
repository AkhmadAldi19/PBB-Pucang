<?php
class petugasModel extends CI_Model{
    
    
    public function fetchAll(){
        $query = $this->db->query("SELECT * FROM admin a
        ORDER BY a.id ");

        return $query;

    }

    public function fetchBelumBayar(){
        $query = $this->db->query("SELECT * FROM `admin` where ket = 'Belum Bayar'");

        return $query;
    }

    public function fetchData($limit, $start){
        $query = $this->db->query("SELECT * FROM admin a
        ORDER BY a.id 
        LIMIT " . $start . ", " . $limit);

        return $query;
    }
    
    public function fetchSingle($id){
        return $this->db->get_where('admin', array('id' => $id));
    }

    public function insertData($data){
        $this->db->insert('admin', $data);
    }

    public function updateData($data){
        $this->db->where("id", $data['id']);
        $this->db->update("admin", $data);
    }
    function deleteData($id){
        $this->db->where('id', $id);
        $this->db->delete('admin');
    }

    public function insertDataBayar($data){
        $this->db->insert('admin', $data);
    }
}
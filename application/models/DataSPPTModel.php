<?php
class DataSPPTModel extends CI_Model{
    public function fetchAll(){
        $query = $this->db->query("SELECT * FROM data_pbb a
        ORDER BY a.id DESC");

        return $query;

    }

    public function fetchBelumBayar(){
        $query = $this->db->query("SELECT * FROM `data_pbb` where ket = 'Belum Bayar'");

        return $query;
    }

    public function fetchData($limit, $start){
        $query = $this->db->query("SELECT * FROM data_pbb a
        ORDER BY a.id DESC
        LIMIT " . $start . ", " . $limit);

        return $query;
    }
    
    public function fetchSingle($id){
        return $this->db->get_where('data_pbb', array('id' => $id));
    }

    // public function insertData($data){
    //     $this->db->insert('data_pbb', $data);
    // }

    public function insertData($data){
        $this->db->insert('data_pbb', $data);
    }

    public function updateData($data){
        $this->db->where("id", $data['id']);
        $this->db->update("data_pbb", $data);
    }
    function deleteData($id){
        $this->db->where('id', $id);
        $this->db->delete('data_pbb');
    }

    public function insertDataBayar($data){
        $this->db->insert('data_bayar', $data);
    }
}
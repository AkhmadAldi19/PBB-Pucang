<?php
class PBBTerhutangModel extends CI_Model{
    public function fetchData($limit, $start){
        $query = $this->db->query("SELECT * FROM data_pbb a
        WHERE ket = 'Belum bayar'
        ORDER BY a.id DESC
        LIMIT " . $start . ", " . $limit);

        return $query;
    }
    
    public function fetchSingle($id){
        return $this->db->get_where('data_pbb', array('id' => $id));
    }

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
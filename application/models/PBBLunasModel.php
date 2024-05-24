<?php
class PBBLunasModel extends CI_Model{
    public function fetchData($limit, $start, $id){
        $query = $this->db->query("SELECT * FROM data_bayar a
        WHERE a.ket = 'Lunas' AND a.id_petugas = $id
        ORDER BY a.id DESC
        LIMIT " . $start . ", " . $limit);

        return $query;
    }
    
}
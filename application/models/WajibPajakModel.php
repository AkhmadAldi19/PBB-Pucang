<?php
class WajibPajakModel extends CI_Model{
    public function fetchData($limit, $start){
        $query = $this->db->query("SELECT * FROM nwp a
        ORDER BY a.id DESC
        LIMIT " . $start . ", " . $limit);

        return $query;
    }
}
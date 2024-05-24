<?php
class SetorPBBModel extends CI_Model{
    public function fetchData($limit, $start, $id){
        $query = $this->db->query("SELECT a.*, b.nama FROM data_setor a
        left outer join data_pbb b ON a.nop = b.nop
        where id_admin = ". $id ."
        ORDER BY a.id DESC
        LIMIT " . $start . ", " . $limit);

        return $query;
    }
    public function fetchSetor($id, $tgl_awal, $tgl_akhir){
        $query = $this->db->query("SELECT a.*, b.nama FROM data_setor a
        left outer join data_pbb b ON a.nop = b.nop
        where id_admin = ". $id ." AND tanggal BETWEEN '". $tgl_awal ."' AND '". $tgl_akhir ."' ORDER BY a.id DESC ");

        return $query;
    }
}
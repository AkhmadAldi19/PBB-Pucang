<?php
class CetakDokumenModel extends CI_Model{

    public function fetchAll(){
        $query = $this->db->query("SELECT * FROM data_pbb a
        WHERE ket = 'Belum bayar'
        ORDER BY a.id DESC");

        return $query;

    }
    public function fetchAllLunas($id){
        $query = $this->db->query("SELECT * FROM data_bayar a
        WHERE a.ket = 'Lunas' AND a.id_petugas = $id
        ORDER BY a.id DESC");

        return $query;
    }

    public function fetchSetor($id, $tgl_awal, $tgl_akhir){
        $query = $this->db->query("SELECT a.*, b.nama FROM data_setor a
        left outer join data_pbb b ON a.nop = b.nop
        where id_admin = ". $id ." AND tanggal BETWEEN '". $tgl_awal ."' AND '". $tgl_akhir ."' ORDER BY a.id DESC ");

        return $query;
    }

    public function fetchPBB($id){
        $query = $this->db->query("SELECT a.*, b.tanggal FROM `data_pbb` a
        LEFT OUTER JOIN data_bayar b ON a.nop = b.nop
        WHERE a.id = " . $id);

        return $query;
    }
}
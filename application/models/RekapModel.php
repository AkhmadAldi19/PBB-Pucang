<?php
class RekapModel extends CI_Model{
    public function fetchData(){
        $query = $this->db->query("SELECT (SELECT count(*) FROM `data_pbb`) as jumlah_nop,
        (SELECT sum(pajak) FROM `data_pbb`) as total_pajak,
        (SELECT count(*) FROM `data_pbb` where ket = 'Lunas') as lunas,
        (SELECT sum(pajak) FROM `data_pbb` where ket = 'Lunas') as pajak_lunas,
        (SELECT count(*) FROM `data_pbb` where ket = 'Belum Bayar') as terhutang,
        (SELECT sum(pajak) FROM `data_pbb` where ket = 'Belum Bayar') as pajak_terhutang,
        (SELECT count(*) FROM `data_pbb` where ket = 'Lunas')/(SELECT count(*) FROM `data_pbb`)*100 as presentase");

        return $query;
    }
     public function fetchSingle(){
        return $this->db->get_where('profil', array('id' => 1));
    }
    
}
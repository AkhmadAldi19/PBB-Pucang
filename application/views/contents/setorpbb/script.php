<script>
$(document).ready(function(){
    var href = "<?=base_url()?>cetakdokumen/pbbsetor/$uri1/$uri2"
    $('#tgl_awal').change(function(){
        href = href.replace('$uri1', $(this).val());
        $('#cetaksetor').attr('href', href);
    });
    $('#tgl_akhir').change(function(){
        href = href.replace('$uri2', $(this).val());
        $('#cetaksetor').attr('href', href);
    });
});
</script>
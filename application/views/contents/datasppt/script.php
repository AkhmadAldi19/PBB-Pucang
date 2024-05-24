<script>
    $(document).ready(function(){
        $('#select').on('change', function() {
            window.location.href= '<?=base_url()?>dataSPPT/bayarPajak/' +this.value;
        });
    });
</script>
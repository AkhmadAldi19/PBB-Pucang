<script>
    $(document).ready(function(){
        $('#allcheck').change(function(){
            if(this.checked){
                $('.check').prop('checked', true);
            }else{
                $('.check').prop('checked', false);
            }
        });
    });
</script>
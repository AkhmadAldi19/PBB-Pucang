    <!-- Jquery JS-->
    <script src="<?=site_url()?>assets/vendor/jquery-3.2.1.min.js"></script>
    <!-- Bootstrap JS-->
    <script src="<?=site_url()?>assets/vendor/bootstrap-4.1/popper.min.js"></script>
    <script src="<?=site_url()?>assets/vendor/bootstrap-4.1/bootstrap.min.js"></script>
    <!-- Latest compiled and minified JavaScript -->
    <script src="<?=site_url()?>assets/vendor/bootstrap-4.1/bootstrap-select.min.js"></script>

    <!-- Vendor JS       -->
    <script src="<?=site_url()?>assets/vendor/slick/slick.min.js">
    </script>
    <script src="<?=site_url()?>assets/vendor/wow/wow.min.js"></script>
    <script src="<?=site_url()?>assets/vendor/animsition/animsition.min.js"></script>
    <script src="<?=site_url()?>assets/vendor/bootstrap-progressbar/bootstrap-progressbar.min.js">
    </script>
    <script src="<?=site_url()?>assets/vendor/counter-up/jquery.waypoints.min.js"></script>
    <script src="<?=site_url()?>assets/vendor/counter-up/jquery.counterup.min.js">
    </script>
    <script src="<?=site_url()?>assets/vendor/circle-progress/circle-progress.min.js"></script>
    <script src="<?=site_url()?>assets/vendor/perfect-scrollbar/perfect-scrollbar.js"></script>
    <script src="<?=site_url()?>assets/vendor/chartjs/Chart.bundle.min.js"></script>
    <script src="<?=site_url()?>assets/vendor/select2/select2.min.js">
    </script>

    <!-- Main JS-->
    <script src="<?=site_url()?>assets/js/main.js"></script>
    
    <script>
      $(document).ready(function(){

        $("#myInput").on("keyup", function() {
          var value = $(this).val().toLowerCase();
          $("#myTable tbody tr").filter(function() {
            $(this).toggle($(this).text().toLowerCase().indexOf(value) > -1)
          });
        });
        $(function () {
          $('[data-toggle="tooltip"]').tooltip()
        });

        $('')
        
      });
    </script>
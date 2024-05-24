<div class="row form-group">
    <div class="col col-md-3">
        <label for="text-input" class=" form-control-label">User Name</label>
    </div>
    <div class="col-12 col-md-9">
        <input type="text"  name="uname" class="form-control" value="<?=!empty($d->uname)?$d->uname:''?>">
    </div>
</div>

<div class="row form-group">
    <div class="col col-md-3">
        <label for="text-input" class=" form-control-label">Nomor HP</label>
    </div>
    <div class="col-12 col-md-9">
        <input type="number"  name="hp" class="form-control" value="<?=!empty($d->hp)?$d->hp:''?>">
    </div>
</div>

<div class="row form-group">
    <div class="col col-md-3">
        <label for="text-input" class=" form-control-label">Kata Sandi</label>
    </div>
    <div class="col-12 col-md-9">
        <input type="password"  name="pass" class="form-control" >
    </div>
</div>

<div class="row form-group">
    <div class="col col-md-3">
        <label for="text-input" class=" form-control-label">Level</label>
    </div>
    <div class="col-12 col-md-9">
        <div class="form-group">
          <select name="level" class="form-control">
                <option><?=!empty($d->level)?$d->level:''?></option>
                <option value="admin">Admin</option>
                <option value="petugas">Petugas</option>
            </select>
                        
        </div>
    </div>
    <input type="hidden"  name="foto" class="form-control" value="1.jpg">
</div>

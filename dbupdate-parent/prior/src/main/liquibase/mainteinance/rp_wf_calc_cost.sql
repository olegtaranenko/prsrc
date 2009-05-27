-- получить текущую себестоимость ненклатуры из базы комтеха
call build_rp_procedure (
	  'stime'
	, 'wf_calc_cost'
	,   'out out_ret float'
	  + ', out out_has_naklad integer'
	  + ', p_id_inv integer'
);


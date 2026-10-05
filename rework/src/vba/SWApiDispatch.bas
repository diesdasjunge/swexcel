Attribute VB_Name = "SWApiDispatchModule"
Option Explicit
Option Private Module

Public Sub SWApiDispatch(ByVal functionName As String, ByVal values As Variant, ByRef result As SWApiResult)
    Select Case functionName
        Case "swe_azalt"
            SWApiCall0 result, values
        Case "swe_azalt_rev"
            SWApiCall1 result, values
        Case "swe_calc"
            SWApiCall2 result, values
        Case "swe_calc_pctr"
            SWApiCall3 result, values
        Case "swe_calc_ut"
            SWApiCall4 result, values
        Case "swe_close"
            SWApiCall5 result, values
        Case "swe_cotrans"
            SWApiCall6 result, values
        Case "swe_cotrans_sp"
            SWApiCall7 result, values
        Case "swe_cs2degstr"
            SWApiCall8 result, values
        Case "swe_cs2lonlatstr"
            SWApiCall9 result, values
        Case "swe_cs2timestr"
            SWApiCall10 result, values
        Case "swe_csnorm"
            SWApiCall11 result, values
        Case "swe_csroundsec"
            SWApiCall12 result, values
        Case "swe_d2l"
            SWApiCall13 result, values
        Case "swe_date_conversion"
            SWApiCall14 result, values
        Case "swe_day_of_week"
            SWApiCall15 result, values
        Case "swe_deg_midp"
            SWApiCall16 result, values
        Case "swe_degnorm"
            SWApiCall17 result, values
        Case "swe_deltat"
            SWApiCall18 result, values
        Case "swe_deltat_ex"
            SWApiCall19 result, values
        Case "swe_difcs2n"
            SWApiCall20 result, values
        Case "swe_difcsn"
            SWApiCall21 result, values
        Case "swe_difdeg2n"
            SWApiCall22 result, values
        Case "swe_difdegn"
            SWApiCall23 result, values
        Case "swe_difrad2n"
            SWApiCall24 result, values
        Case "swe_fixstar"
            SWApiCall25 result, values
        Case "swe_fixstar2"
            SWApiCall26 result, values
        Case "swe_fixstar2_mag"
            SWApiCall27 result, values
        Case "swe_fixstar2_ut"
            SWApiCall28 result, values
        Case "swe_fixstar_mag"
            SWApiCall29 result, values
        Case "swe_fixstar_ut"
            SWApiCall30 result, values
        Case "swe_gauquelin_sector"
            SWApiCall31 result, values
        Case "swe_get_astro_models"
            SWApiCall32 result, values
        Case "swe_get_ayanamsa"
            SWApiCall33 result, values
        Case "swe_get_ayanamsa_ex"
            SWApiCall34 result, values
        Case "swe_get_ayanamsa_ex_ut"
            SWApiCall35 result, values
        Case "swe_get_ayanamsa_name"
            SWApiCall36 result, values
        Case "swe_get_ayanamsa_ut"
            SWApiCall37 result, values
        Case "swe_get_current_file_data"
            SWApiCall38 result, values
        Case "swe_get_library_path"
            SWApiCall39 result, values
        Case "swe_get_orbital_elements"
            SWApiCall40 result, values
        Case "swe_get_planet_name"
            SWApiCall41 result, values
        Case "swe_get_tid_acc"
            SWApiCall42 result, values
        Case "swe_heliacal_angle"
            SWApiCall43 result, values
        Case "swe_heliacal_pheno_ut"
            SWApiCall44 result, values
        Case "swe_heliacal_ut"
            SWApiCall45 result, values
        Case "swe_helio_cross"
            SWApiCall46 result, values
        Case "swe_helio_cross_ut"
            SWApiCall47 result, values
        Case "swe_house_name"
            SWApiCall48 result, values
        Case "swe_house_pos"
            SWApiCall49 result, values
        Case "swe_houses"
            SWApiCall50 result, values
        Case "swe_houses_armc"
            SWApiCall51 result, values
        Case "swe_houses_armc_ex2"
            SWApiCall52 result, values
        Case "swe_houses_ex"
            SWApiCall53 result, values
        Case "swe_houses_ex2"
            SWApiCall54 result, values
        Case "swe_jdet_to_utc"
            SWApiCall55 result, values
        Case "swe_jdut1_to_utc"
            SWApiCall56 result, values
        Case "swe_julday"
            SWApiCall57 result, values
        Case "swe_lat_to_lmt"
            SWApiCall58 result, values
        Case "swe_lmt_to_lat"
            SWApiCall59 result, values
        Case "swe_lun_eclipse_how"
            SWApiCall60 result, values
        Case "swe_lun_eclipse_when"
            SWApiCall61 result, values
        Case "swe_lun_eclipse_when_loc"
            SWApiCall62 result, values
        Case "swe_lun_occult_when_glob"
            SWApiCall63 result, values
        Case "swe_lun_occult_when_loc"
            SWApiCall64 result, values
        Case "swe_lun_occult_where"
            SWApiCall65 result, values
        Case "swe_mooncross"
            SWApiCall66 result, values
        Case "swe_mooncross_node"
            SWApiCall67 result, values
        Case "swe_mooncross_node_ut"
            SWApiCall68 result, values
        Case "swe_mooncross_ut"
            SWApiCall69 result, values
        Case "swe_nod_aps"
            SWApiCall70 result, values
        Case "swe_nod_aps_ut"
            SWApiCall71 result, values
        Case "swe_orbit_max_min_true_distance"
            SWApiCall72 result, values
        Case "swe_pheno"
            SWApiCall73 result, values
        Case "swe_pheno_ut"
            SWApiCall74 result, values
        Case "swe_rad_midp"
            SWApiCall75 result, values
        Case "swe_radnorm"
            SWApiCall76 result, values
        Case "swe_refrac"
            SWApiCall77 result, values
        Case "swe_refrac_extended"
            SWApiCall78 result, values
        Case "swe_revjul"
            SWApiCall79 result, values
        Case "swe_rise_trans"
            SWApiCall80 result, values
        Case "swe_rise_trans_true_hor"
            SWApiCall81 result, values
        Case "swe_set_astro_models"
            SWApiCall82 result, values
        Case "swe_set_delta_t_userdef"
            SWApiCall83 result, values
        Case "swe_set_ephe_path"
            SWApiCall84 result, values
        Case "swe_set_interpolate_nut"
            SWApiCall85 result, values
        Case "swe_set_jpl_file"
            SWApiCall86 result, values
        Case "swe_set_lapse_rate"
            SWApiCall87 result, values
        Case "swe_set_sid_mode"
            SWApiCall88 result, values
        Case "swe_set_tid_acc"
            SWApiCall89 result, values
        Case "swe_set_topo"
            SWApiCall90 result, values
        Case "swe_sidtime"
            SWApiCall91 result, values
        Case "swe_sidtime0"
            SWApiCall92 result, values
        Case "swe_sol_eclipse_how"
            SWApiCall93 result, values
        Case "swe_sol_eclipse_when_glob"
            SWApiCall94 result, values
        Case "swe_sol_eclipse_when_loc"
            SWApiCall95 result, values
        Case "swe_sol_eclipse_where"
            SWApiCall96 result, values
        Case "swe_solcross"
            SWApiCall97 result, values
        Case "swe_solcross_ut"
            SWApiCall98 result, values
        Case "swe_split_deg"
            SWApiCall99 result, values
        Case "swe_time_equ"
            SWApiCall100 result, values
        Case "swe_topo_arcus_visionis"
            SWApiCall101 result, values
        Case "swe_utc_time_zone"
            SWApiCall102 result, values
        Case "swe_utc_to_jd"
            SWApiCall103 result, values
        Case "swe_version"
            SWApiCall104 result, values
        Case "swe_vis_limit_mag"
            SWApiCall105 result, values
        Case Else
            SWRaise "Unknown Swiss Ephemeris function."
    End Select
End Sub

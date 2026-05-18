Attribute VB_Name = "Module2"
'*************************************************************************************************************
'
' This module contains code for the multi-species growth model.
'
'*************************************************************************************************************

Option Explicit
    
Public H30 As Double, TBH As Double, D300_30_est As Double, rotlength As Double, _
    T1 As Double, H1 As Double, N1 As Double, D1 As Double, T2 As Double, H2 As Double, mode As Integer, DiaDist As Integer, _
    N(200) As Double, MTH(200) As Double, DBH(200) As Double, BA(200) As Double, Vol(200) As Double, T_adj(200) As Double, WoodDen(200) As Double, _
    Stock_hist_T(6) As Integer, Stock_hist_N(6) As Double, Stock_hist_thin_coeff(6) As Double, Nthins As Integer, nstems As Long, _
    Treelist(10000, 6) As Double, petA As Double, petB As Double, k_mort, weibull_b As Double, weibull_a As Double, _
    bh_mod As Double, Error_flag As Boolean, Pre_thin_vol(4) As Double, Pre_thin_N(4) As Double, Pre_thin_dbh(4) As Double, _
    Thin_dbh(4) As Double, Stemlist(10000, 6) As Double, stemno As Long, stem_stand_vol_adj As Double, Thin_vol(4) As Double, _
    WoodDensity_Adjustment As Double, CWD_half_life As Double, Carbon_method As Integer, Volume_function As Integer, weibull_CV As Double, _
    alpha0 As Double, alpha1 As Double, alpha2 As Double, beta1 As Double, beta2 As Double, beta3 As Double, beta4 As Double, _
    beta5 As Double, Stock_hist_Type(5) As Integer, log_length As Double, u As Double, V As Double, w As Double, _
    min_SED As Double, AGCWD_half_life As Double, BGCWD_half_life As Double, log_losses As Double, harvest_sum(30, 11) As Double, _
    Harvest_volume As Double, log_volume(2000) As Double, logno As Long, Logs(10000, 10) As Double, Cali As Boolean, _
    break_height As Double, Species As String, prune_age(4) As Double, prune_N(4) As Double, prune_height(4) As Double, _
    latitude As Double, elevation As Double, Soil_C As Double, Soil_N As Double, MAT As Double, drift As Double, _
    MTH_model As String, MTH_form As String, MTH_a As Double, MTH_b As Double, MTH_c As Double, _
    DBH_model As String, DBH_form As String, DBH_a As Double, DBH_b As Double, DBH_c As Double, DBH_d As Double, DBH_f As Double, _
    DBH_g As Double, DBH_h As Double, DBH_k As Double, MTH_MnHt_a   As Double, MTH_MnHt_b As Double, _
    VOL_type As Double, VOL_u As Double, VOL_v As Double, VOL_w As Double, VOL_z As Double, THINCOEF As Double, _
    MORT_k As Double, MORT_m As Double, MORT_n As Double, Den_a As Double, Den_b As Double, thin_age(4) As Double, _
    PRUNEHT As Double, Check_errors As Boolean, Minimal_run As Boolean
    
' MTH model parameters
Const MTH_model_red As String = "Korf"
Const MTH_form_red As String = "GADA"
Const MTH_a_red As Double = 0
Const MTH_b_red As Double = 40.82
Const MTH_c_red As Double = 0.4619
Const MTH_model_lus As String = "Richards"
Const MTH_form_lus As String = "CA"
Const MTH_a_lus As Double = 38.79
Const MTH_b_lus As Double = 0
Const MTH_c_lus As Double = 1.099
Const MTH_model_mac As String = "Richards"
Const MTH_form_mac As String = "CA"
Const MTH_a_mac As Double = 47.37
Const MTH_b_mac As Double = 0
Const MTH_c_mac As Double = 1.062
Const MTH_model_bla As String = "Richards"
Const MTH_form_bla As String = "Anamorphic"
Const MTH_a_bla As Double = 0
Const MTH_b_bla As Double = 0.01325
Const MTH_c_bla As Double = 0.9302
Const MTH_model_reg As String = "Korf"
Const MTH_form_reg As String = "CA"
Const MTH_a_reg As Double = 259.8
Const MTH_b_reg As Double = 0
Const MTH_c_reg As Double = 0.3387
Const MTH_model_fas As String = "Hossfeld"
Const MTH_form_fas As String = "CA"
Const MTH_a_fas As Double = 62.71
Const MTH_b_fas As Double = 0
Const MTH_c_fas As Double = 1.392
Const MTH_model_nit As String = "Korf"
Const MTH_form_nit As String = "CA"
Const MTH_a_nit As Double = 116.04
Const MTH_b_nit As Double = 0
Const MTH_c_nit As Double = 0.4539
Const MTH_model_del As String = "Hossfeld"
Const MTH_form_del As String = "GADA"
Const MTH_a_del As Double = 49.93
Const MTH_b_del As Double = 829
Const MTH_c_del As Double = 1.201
Const MTH_model_sal As String = "Korf"
Const MTH_form_sal As String = "CA"
Const MTH_a_sal As Double = 106.03
Const MTH_b_sal As Double = 0
Const MTH_c_sal As Double = 0.4586

' DBH model parameters
Const DBH_model_red As String = "Korf"
Const DBH_form_red As String = "Anamorphic"
Const DBH_a_red As Double = 0
Const DBH_b_red As Double = 5.629
Const DBH_c_red As Double = 0.44
Const DBH_model_lus As String = "Korf"
Const DBH_form_lus As String = "CA"
Const DBH_a_lus As Double = 83.28
Const DBH_b_lus As Double = 0
Const DBH_c_lus As Double = 0.6974
Const DBH_model_mac As String = "Korf"
Const DBH_form_mac As String = "CA"
Const DBH_a_mac As Double = 83.14
Const DBH_b_mac As Double = 0
Const DBH_c_mac As Double = 0.71
Const DBH_model_bla As String = "Richards"
Const DBH_form_bla As String = "Anamorphic"
Const DBH_a_bla As Double = 0
Const DBH_b_bla As Double = 0.03294
Const DBH_c_bla As Double = 0.9736
Const DBH_model_reg As String = "Korf"
Const DBH_form_reg As String = "Anamorphic"
Const DBH_a_reg As Double = 0
Const DBH_b_reg As Double = 4.298
Const DBH_c_reg As Double = 0.3258
Const DBH_model_fas As String = "Korf"
Const DBH_form_fas As String = "GADA"
Const DBH_a_fas As Double = 0
Const DBH_b_fas As Double = 22.62
Const DBH_c_fas As Double = 0.4484
Const DBH_model_nit As String = "Richards"
Const DBH_form_nit As String = "Anamorphic"
Const DBH_a_nit As Double = 0
Const DBH_b_nit As Double = 0.06808
Const DBH_c_nit As Double = 1.0733
Const DBH_model_del As String = "Korf"
Const DBH_form_del As String = "Anamorphic"
Const DBH_a_del As Double = 0
Const DBH_b_del As Double = 5.037
Const DBH_c_del As Double = 0.3601
Const DBH_model_sal As String = "Korf"
Const DBH_form_sal As String = "GADA"
Const DBH_a_sal As Double = 0
Const DBH_b_sal As Double = 18.13
Const DBH_c_sal As Double = 0.4271

' DBH stand density adjustment model parameters
Const DBH_d_red As Double = 0.5663
Const DBH_f_red As Double = 0.0912
Const DBH_g_red As Double = 39.94
Const DBH_h_red As Double = 0.735
Const DBH_k_red As Double = 0
Const DBH_d_lus As Double = 0.3659
Const DBH_f_lus As Double = 0.4624
Const DBH_g_lus As Double = 32.02
Const DBH_h_lus As Double = 0
Const DBH_k_lus As Double = -2.39
Const DBH_d_mac As Double = 0.4213
Const DBH_f_mac As Double = 0.2645
Const DBH_g_mac As Double = 20.66
Const DBH_h_mac As Double = 0
Const DBH_k_mac As Double = -0.62
Const DBH_d_bla As Double = 0.1994
Const DBH_f_bla As Double = 0.5
Const DBH_g_bla As Double = 10.88
Const DBH_h_bla As Double = 0.0457
Const DBH_k_bla As Double = 0
Const DBH_d_reg As Double = 0.3697
Const DBH_f_reg As Double = 0.199
Const DBH_g_reg As Double = 12.34
Const DBH_h_reg As Double = 0.3998
Const DBH_k_reg As Double = 0
Const DBH_d_fas As Double = 0.405
Const DBH_f_fas As Double = 0.141
Const DBH_g_fas As Double = 16.84
Const DBH_h_fas As Double = 1.0695
Const DBH_k_fas As Double = 0
Const DBH_d_nit As Double = 0.6215
Const DBH_f_nit As Double = 0.184
Const DBH_g_nit As Double = 21.32
Const DBH_h_nit As Double = 0.0558
Const DBH_k_nit As Double = 0
Const DBH_d_del As Double = 0.6215
Const DBH_f_del As Double = 0.168
Const DBH_g_del As Double = 26.38
Const DBH_h_del As Double = 0.7319
Const DBH_k_del As Double = 0
Const DBH_d_sal As Double = 0.3161
Const DBH_f_sal As Double = 0.232
Const DBH_g_sal As Double = 9.6
Const DBH_h_sal As Double = 0.2417
Const DBH_k_sal As Double = 0

' Parameters of models for predicting MTH from mean height and stand density
Const MTH_MnHt_a_red As Double = 0.1376
Const MTH_MnHt_b_red As Double = -0.00622
Const MTH_MnHt_a_lus As Double = 0.1921
Const MTH_MnHt_b_lus As Double = -0.00114
Const MTH_MnHt_a_mac As Double = 0.1921
Const MTH_MnHt_b_mac As Double = -0.00114
Const MTH_MnHt_a_bla As Double = 0.1271
Const MTH_MnHt_b_bla As Double = -0.00285
Const MTH_MnHt_a_reg As Double = 0.1616
Const MTH_MnHt_b_reg As Double = -0.00464
Const MTH_MnHt_a_fas As Double = 0.2059
Const MTH_MnHt_b_fas As Double = -0.00165
Const MTH_MnHt_a_nit As Double = 0.1401
Const MTH_MnHt_b_nit As Double = -0.00262
Const MTH_MnHt_a_del As Double = 0.1793
Const MTH_MnHt_b_del As Double = -0.00278
Const MTH_MnHt_a_sal As Double = 0.2146
Const MTH_MnHt_b_sal As Double = -0.00164
Const MTH_MnHt_a_rad As Double = 0.07
Const MTH_MnHt_b_rad As Double = -0.00399
Const MTH_MnHt_a_dfr As Double = 0.106
Const MTH_MnHt_b_dfr As Double = -0.228

' Parameters of stand-level volume functions
Const VOL_type_red1 As Double = 1
Const VOL_u_red1 As Double = 0
Const VOL_v_red1 As Double = 0.4872
Const VOL_w_red1 As Double = 0.1796
Const VOL_z_red1 As Double = 0
Const VOL_type_red2 As Double = 1
Const VOL_u_red2 As Double = 0.2251
Const VOL_v_red2 As Double = 0.7221
Const VOL_w_red2 As Double = 0.8593
Const VOL_z_red2 As Double = 0
Const VOL_type_lus As Double = 1
Const VOL_u_lus As Double = 0.2944
Const VOL_v_lus As Double = 0.3973
Const VOL_w_lus As Double = 0.5866
Const VOL_z_lus As Double = 0
Const VOL_type_mac As Double = 1
Const VOL_u_mac As Double = 0.2944
Const VOL_v_mac As Double = 0.3973
Const VOL_w_mac As Double = 0.5866
Const VOL_z_mac As Double = 0
Const VOL_type_bla As Double = 2
Const VOL_u_bla As Double = -0.2801
Const VOL_v_bla As Double = 0.9704
Const VOL_w_bla As Double = 0.5543
Const VOL_z_bla As Double = 0.079
Const VOL_type_reg As Double = 2
Const VOL_u_reg As Double = -1.0996
Const VOL_v_reg As Double = 0.9602
Const VOL_w_reg As Double = 1.0526
Const VOL_z_reg As Double = 0
Const VOL_type_fas As Double = 2
Const VOL_u_fas As Double = -0.9468
Const VOL_v_fas As Double = 0.9892
Const VOL_w_fas As Double = 0.8995
Const VOL_z_fas As Double = 0
Const VOL_type_nit As Double = 2
Const VOL_u_nit As Double = 0.012
Const VOL_v_nit As Double = 0.9653
Const VOL_w_nit As Double = 0.4611
Const VOL_z_nit As Double = 0.0716
Const VOL_type_del As Double = 2
Const VOL_u_del As Double = -0.3391
Const VOL_v_del As Double = 0.9541
Const VOL_w_del As Double = 0.8369
Const VOL_z_del As Double = 0
Const VOL_type_sal As Double = 2
Const VOL_u_sal As Double = 0.2173
Const VOL_v_sal As Double = 0.9632
Const VOL_w_sal As Double = 0.3116
Const VOL_z_sal As Double = 0.0972
Const VOL_type_rad As Double = 1    ' Kimberley & Beets, 2007
Const VOL_u_rad As Double = 0.317
Const VOL_v_rad As Double = 0.942
Const VOL_w_rad As Double = 1.161
Const VOL_z_rad As Double = 0
Const VOL_type_dfr As Double = 3    ' Kimberley & Beets, 2007
Const VOL_u_dfr As Double = 0.3208
Const VOL_v_dfr As Double = 0.928
Const VOL_w_dfr As Double = 0
Const VOL_z_dfr As Double = 0

' Thinning coefficients
Const THINCOEF_red As Double = 0.78
Const THINCOEF_cyp As Double = 0.7
Const THINCOEF_bla As Double = 0.765
Const THINCOEF_euc As Double = 0.7
Const THINCOEF_rad As Double = 0.784
Const THINCOEF_dfr As Double = 0.705

' Mortality function parameters
Const MORT_k_red As Double = 0.00167
Const MORT_m_red As Double = 0.00368
Const MORT_n_red As Double = 3.56
Const MORT_k_lus_NI As Double = 0.0049
Const MORT_m_lus_NI As Double = 0.0314
Const MORT_n_lus_NI As Double = 1.41
Const MORT_k_lus_SI As Double = 0.0054
Const MORT_m_lus_SI As Double = 0.0314
Const MORT_n_lus_SI As Double = 1.41
Const MORT_k_mac_NI As Double = 0.0103
Const MORT_m_mac_NI As Double = 0.0314
Const MORT_n_mac_NI As Double = 1.41
Const MORT_k_mac_SI As Double = 0
Const MORT_m_mac_SI As Double = 0.0314
Const MORT_n_mac_SI As Double = 1.41
Const MORT_k_bla As Double = 0.00956
Const MORT_m_bla As Double = 0.03249
Const MORT_n_bla As Double = 2.798
Const MORT_k_reg As Double = 0.0159
Const MORT_m_reg As Double = 0.03249
Const MORT_n_reg As Double = 2.798
Const MORT_k_fas As Double = 0.0151
Const MORT_m_fas As Double = 0.03249
Const MORT_n_fas As Double = 2.798
Const MORT_k_nit_NI As Double = 0.0183
Const MORT_m_nit_NI As Double = 0.4692
Const MORT_n_nit_NI As Double = 2.798
Const MORT_k_nit_SI As Double = 0.00608
Const MORT_m_nit_SI As Double = 0.03249
Const MORT_n_nit_SI As Double = 2.798
Const MORT_k_del As Double = 0.0142
Const MORT_m_del As Double = 0.03249
Const MORT_n_del As Double = 2.798
Const MORT_k_sal As Double = 0.00742
Const MORT_m_sal As Double = 0.03249
Const MORT_n_sal As Double = 2.798

' Parameters of wood density models
Const DEN_a_red As Double = 339
Const DEN_b_red As Double = 0
Const DEN_a_cyp As Double = 404
Const DEN_b_cyp As Double = 0
Const DEN_a_bla As Double = 321
Const DEN_b_bla As Double = 60.7
Const DEN_a_reg As Double = 313
Const DEN_b_reg As Double = 41.1
Const DEN_a_fas As Double = 351
Const DEN_b_fas As Double = 43.2
Const DEN_a_nit As Double = 313
Const DEN_b_nit As Double = 60.7
Const DEN_a_del As Double = 318
Const DEN_b_del As Double = 41.1
Const DEN_a_sal As Double = 320
Const DEN_b_sal As Double = 78.2


' T458 redwood volume and taper model parameters
Const alpha0_458 = 0.702
Const alpha1_458 = 0.5646
Const alpha2_458 = -0.6188
Const beta1_458 = 2.6295
Const beta2_458 = 0.1406
Const beta3_458 = 0.1455
Const beta4_458 = -0.1275
Const beta5_458 = 22.7873

' T472 redwood volume and taper model parameters
Const alpha0_472 = 0.7472
Const alpha1_472 = 0.361
Const alpha2_472 = -0.4524
Const beta1_472 = 3.8508
Const beta2_472 = 0.2725
Const beta3_472 = 0.2658
Const beta4_472 = 0.0509
Const beta5_472 = 21.9598

' Other constants
Public Const bh = 1.4
Public Const PI = 3.14159265359873

' Parameters of redwood diamater under-bark model (from T468 report)
Const uba1 = 0.7468
Const uba2 = 0.3609
Const uba3 = -0.4524

' Parameters of redwood stand-level volume function T472
Const uv472 = 0.28606
Const vv472 = 5.6238
Const wv472 = 2.4217

' Parameters of cypress stand-level volume function
Const uvCYP = 0.2944
Const vvCYP = 0.3973
Const wvCYP = 0.5866

' Redwood height conversion equation parameters
Const rcRED = 0.1376
Const scRED = -0.006221

' Cypress height conversion equation parameters
Const rcCYP = 0.1921
Const scCYP = -0.00114

' Crown height model parameters
Const pCH = -39.1
Const qch = 0.78
Const rch = 4.87

Function MTH_mod(T As Double, SI As Double, MTH_model As String, MTH_form As String, MTH_a As Double, MTH_b As Double, MTH_c As Double)
    ' Predict MTH from age T, Site Index SI
    MTH_mod = 0.3 + Y(T, MTH_model, MTH_form, SI - 0.3, 30, MTH_a, MTH_b, MTH_c)
End Function

Function SI_eqn(T As Double, MTH As Double, MTH_model As String, MTH_form As String, MTH_a As Double, MTH_b As Double, MTH_c As Double)
    ' Predict SI from MTH at age T
    SI_eqn = 0.3 + Y0(30, MTH_model, MTH_form, MTH - 0.3, T, MTH_a, MTH_b, MTH_c)
End Function

Function AgeBH(T As Double, H As Double, MTH_model As String, MTH_form As String, MTH_a As Double, MTH_b As Double, MTH_c As Double)
    ' Determine when stand achieves 1.4 m height from a measurement of height H at age t years with MTH model b parameter = bh_mod
    Dim R As Double, ah As Double
    AgeBH = TP(1.4 - 0.3, MTH_model, MTH_form, H - 0.3, T, MTH_a, MTH_b, MTH_c)
End Function

Function D300(T As Double, TBH As Double, D300_30 As Double, DBH_model As String, DBH_form As String, DBH_a As Double, DBH_b As Double, DBH_c As Double)
    ' Calculate D300 from age, breast height age and D300_30
    D300 = Y(T - TBH, DBH_model, DBH_form, D300_30, 30 - TBH, DBH_a, DBH_b, DBH_c)
End Function

Function DBH_mod(T As Double, D300_30 As Double, SI As Double, TBH As Double, N As Double, DBH_model As String, _
        DBH_form As String, DBH_a As Double, DBH_b As Double, DBH_c As Double, DBH_d As Double, DBH_f As Double, DBH_g As Double, _
        DBH_h As Double, DBH_k As Double, MTH_model, MTH_form, MTH_a, MTH_b, MTH_c)
    ' Calculate DBH from age, D300_30, SI, TBH and stocking
    Dim D300_est As Double, Ntemp As Double
    D300_est = D300(T, TBH, D300_30, DBH_model, DBH_form, DBH_a, DBH_b, DBH_c)
    Ntemp = N
    If Ntemp > 200 * Exp(1 / DBH_d) Then Ntemp = 200 * Exp(1 / DBH_d)
    DBH_mod = D300_est - DBH_d / DBH_f * (Log(Ntemp) - Log(300)) * Log(1 + Exp(DBH_f * (D300_est - (DBH_g + DBH_h * (SI - 30) _
        + DBH_k * Log(Ntemp)))))
End Function

Function Vol_stand(MTH As Double, DBH As Double, N As Double, VOL_type As Double, VOL_u As Double, VOL_v As Double, VOL_w As Double, VOL_z As Double)
    ' Stand-level volume function
    Dim BA As Double, u As Double, V As Double, w As Double
    BA = N * PI * (DBH / 200) ^ 2
    If DBH <= 0 Or MTH <= 1.4 Then
        Vol_stand = 0
    ElseIf VOL_type = 1 Then
        Vol_stand = MTH * BA * (VOL_v * (MTH - 1.4) ^ (-VOL_w) + VOL_u)
    ElseIf VOL_type = 2 Then
        Vol_stand = Exp(VOL_u + VOL_v * Log(BA) + VOL_w * Log(MTH) + VOL_z * (Log(MTH)) ^ 2)
    ElseIf VOL_type = 3 Then
        Vol_stand = BA * (VOL_v + VOL_u * MTH)
    End If
End Function

Public Function MnHt_from_MTH(MTH As Double, N As Double, MTH_MnHt_a As Double, MTH_MnHt_b As Double)
    ' Estimate mean height from MTH and stocking
    MnHt_from_MTH = MTH * (1 - MTH_MnHt_a * (1 - Exp(MTH_MnHt_b * (N - 100))))
End Function

Public Function MTH_from_MnHt(MeanHeight As Double, N As Double, MTH_MnHt_a As Double, MTH_MnHt_b As Double)
    ' Estimate MTH from mean height and stocking
    MTH_from_MnHt = MeanHeight / (1 - MTH_MnHt_a * (1 - Exp(MTH_MnHt_b * (N - 100))))
End Function

Function D300_30_from_I300_SI(I300 As Double, SI As Double, VOL_type As Double, VOL_u As Double, VOL_v As Double, VOL_w As Double, VOL_z As Double)
    ' Calculate D300_30 from 300 Index and SI
    Dim BA As Double
    If VOL_type = 1 Then
        BA = (I300 * 30) / (SI * (VOL_v * (SI - 1.4) ^ (-VOL_w) + VOL_u))
    Else
        BA = Exp((Log(I300 * 30) - VOL_u - VOL_w * Log(SI) - VOL_z * (Log(SI)) ^ 2) / VOL_v)
    End If
    D300_30_from_I300_SI = 200 * Sqr(BA / 300 / PI)
End Function

Function I300_from_SI_D300_30(SI As Double, D300_30 As Double, VOL_type As Double, VOL_u As Double, VOL_v As Double, VOL_w As Double, VOL_z As Double)
    ' Calculate 300 Index from D300_30 and SI
    I300_from_SI_D300_30 = Vol_stand(SI, D300_30, 300, VOL_type, VOL_u, VOL_v, VOL_w, VOL_z) / 30
End Function

Function N_Mort(N0 As Double, DBH As Double, deltaT As Double, MORT_k As Double, MORT_m As Double, MORT_n As Double)
    Dim sdi As Double, mort As Double, mm As Double, nm As Double
    If DBH <= 0 Then
        mort = 100 * (MORT_k)
    Else
        sdi = (0.405 * N0 * (0.394 * DBH / 10) ^ 1.605) / 1000  'Reinekes SDI divided by 1000
        mort = 100 * (MORT_k + MORT_m * sdi ^ MORT_n)
    End If
    N_Mort = N0 * (1 - mort / 100) ^ deltaT
End Function

Public Sub run_mod()
    Check_errors = True
    Minimal_run = False
    Call run_model
End Sub

Public Sub run_model()
    ' Run Growth model
    
    ' Prevent Computer Screen from running
    Application.ScreenUpdating = True

    Error_flag = False
    Cali = False
    Worksheets("Inputs").Activate
    Call Input_parameters   'Input stocking history and other parameters
    If Check_errors Then Call Error_checks_1 'Check input parameters for errors
    If Error_flag Then Exit Sub
    Call Input_tree_list    ' Input unscaled tree list or generate it using a Weibull distribution
    If mode = 3 Then
        If Check_errors Then Call Error_checks_4 'Check tree list for errors
        If Error_flag Then Exit Sub
        Call Process_tree_list  'Derive stand metrics from tree list
    End If
    If mode = 2 Or mode = 3 Then
        If Check_errors Then Call Error_checks_3 'Check input stand metrics for errors
        If Error_flag Then Exit Sub
        If Species = "Radiata pine" Then
            Call Calibrate_radiata
        ElseIf Species = "Douglas-fir" Then
            Call Calibrate_dfir
        Else
            Call Calibrate
        End If
    End If
    If Check_errors Then Call Error_checks_2 'Check inputs for errors
    If Error_flag Then Exit Sub
    
    'Estimate breast height age
    Call Input_parameters   'Input stocking history and other parameters
    If Species = "Coast redwood" And T2 <> 0 Then MTH_b = MTHmodel_b(30, H30, T2, H2)
    TBH = AgeBH(30, H30, MTH_model, MTH_form, MTH_a, MTH_b, MTH_c)
    
    stemno = 1
    logno = 1
    D300_30_est = D300_30_from_I300_SI(I300, H30, VOL_type, VOL_u, VOL_v, VOL_w, VOL_z)
    If Species = "Radiata pine" Then
        Call Yield_Table_radiata
    ElseIf Species = "Douglas-fir" Then
        Call Yield_Table_dfir
    Else
        Call Yield_Table    'Generate yield table
    End If
    If Check_errors Then Call Error_checks_5
    Call Initialise_C_Change    'Set up C_Change inputs
    Call Mortality_Volume   'Estimate mortality volume
    If Not Minimal_run Then
        If Species <> "Coast redwood" Then
            Call Run_C_Change   'Run C_Change for cypress species
        Else
            Call Kizha_Han  'For redwood, use Kizha Han allometric models, stem volume and density to calculate carbon
        End If
    End If
    Call Output_table
    If DiaDist = 2 Then
        If Check_errors Then Call Error_checks_4b
        If Error_flag Then Exit Sub
    End If
    
    'Allow Computer Screen to refresh
    Worksheets("Growth model").Activate
    Application.ScreenUpdating = True
End Sub

Sub Input_parameters()
    Dim i As Integer
    Worksheets("Inputs").Activate
    Species = Cells(2, 5)
    I300 = Cells(3, 5)
    H30 = Cells(4, 5)
    T1 = Cells(20, 5)   ' Age at calibration
    H1 = Cells(22, 5)   ' Height at calibration
    N1 = Cells(21, 5)   ' Stocking at calibration
    If Cells(22, 6) = 2 Then H1 = MTH_from_MnHt(H1, N1, MTH_MnHt_a, MTH_MnHt_b)   ' If necessary, comvert calibration height from mean height to MTH
    D1 = Cells(23, 5)   ' Calibration BA
    If Cells(23, 6) = 1 Then D1 = 200 * Sqr(D1 / N1 / PI)   ' If necessary calculate calibration DBH from BA
    T2 = Cells(24, 5)
    H2 = Cells(25, 5)

    If Species = "Coast redwood" Then
        MTH_model = MTH_model_red
        MTH_form = MTH_form_red
        MTH_a = MTH_a_red
        MTH_b = MTH_b_red
        MTH_c = MTH_c_red
        DBH_model = DBH_model_red
        DBH_form = DBH_form_red
        DBH_a = DBH_a_red
        DBH_b = DBH_b_red
        DBH_c = DBH_c_red
        DBH_d = DBH_d_red
        DBH_f = DBH_f_red
        DBH_g = DBH_g_red
        DBH_h = DBH_h_red
        DBH_k = DBH_k_red
        MTH_MnHt_a = MTH_MnHt_a_red
        MTH_MnHt_b = MTH_MnHt_b_red
        VOL_type = VOL_type_red2
        VOL_u = VOL_u_red2
        VOL_v = VOL_v_red2
        VOL_w = VOL_w_red2
        VOL_z = VOL_z_red2
        THINCOEF = THINCOEF_red
        MORT_k = MORT_k_red
        MORT_m = MORT_m_red
        MORT_n = MORT_n_red
        Den_a = DEN_a_red
        Den_b = DEN_b_red
    ElseIf Species = "Cupressus macrocarpa (N.I.)" Then
        MTH_model = MTH_model_mac
        MTH_form = MTH_form_mac
        MTH_a = MTH_a_mac
        MTH_b = MTH_b_mac
        MTH_c = MTH_c_mac
        DBH_model = DBH_model_mac
        DBH_form = DBH_form_mac
        DBH_a = DBH_a_mac
        DBH_b = DBH_b_mac
        DBH_c = DBH_c_mac
        DBH_d = DBH_d_mac
        DBH_f = DBH_f_mac
        DBH_g = DBH_g_mac
        DBH_h = DBH_h_mac
        DBH_k = DBH_k_mac
        MTH_MnHt_a = MTH_MnHt_a_mac
        MTH_MnHt_b = MTH_MnHt_b_mac
        VOL_type = VOL_type_mac
        VOL_u = VOL_u_mac
        VOL_v = VOL_v_mac
        VOL_w = VOL_w_mac
        VOL_z = VOL_z_mac
        THINCOEF = THINCOEF_cyp
        MORT_k = MORT_k_mac_NI
        MORT_m = MORT_m_mac_NI
        MORT_n = MORT_n_mac_NI
        Den_a = DEN_a_cyp
        Den_b = DEN_b_cyp
    ElseIf Species = "Cupressus macrocarpa (S.I.)" Then
        MTH_model = MTH_model_mac
        MTH_form = MTH_form_mac
        MTH_a = MTH_a_mac
        MTH_b = MTH_b_mac
        MTH_c = MTH_c_mac
        DBH_model = DBH_model_mac
        DBH_form = DBH_form_mac
        DBH_a = DBH_a_mac
        DBH_b = DBH_b_mac
        DBH_c = DBH_c_mac
        DBH_d = DBH_d_mac
        DBH_f = DBH_f_mac
        DBH_g = DBH_g_mac
        DBH_h = DBH_h_mac
        DBH_k = DBH_k_mac
        MTH_MnHt_a = MTH_MnHt_a_mac
        MTH_MnHt_b = MTH_MnHt_b_mac
        VOL_type = VOL_type_mac
        VOL_u = VOL_u_mac
        VOL_v = VOL_v_mac
        VOL_w = VOL_w_mac
        VOL_z = VOL_z_mac
        THINCOEF = THINCOEF_cyp
        MORT_k = MORT_k_mac_SI
        MORT_m = MORT_m_mac_SI
        MORT_n = MORT_n_mac_SI
        Den_a = DEN_a_cyp
        Den_b = DEN_b_cyp
    ElseIf Species = "Cupressus lusitanica (N.I.)" Then
        MTH_model = MTH_model_lus
        MTH_form = MTH_form_lus
        MTH_a = MTH_a_lus
        MTH_b = MTH_b_lus
        MTH_c = MTH_c_lus
        DBH_model = DBH_model_lus
        DBH_form = DBH_form_lus
        DBH_a = DBH_a_lus
        DBH_b = DBH_b_lus
        DBH_c = DBH_c_lus
        DBH_d = DBH_d_lus
        DBH_f = DBH_f_lus
        DBH_g = DBH_g_lus
        DBH_h = DBH_h_lus
        DBH_k = DBH_k_lus
        MTH_MnHt_a = MTH_MnHt_a_lus
        MTH_MnHt_b = MTH_MnHt_b_lus
        VOL_type = VOL_type_lus
        VOL_u = VOL_u_lus
        VOL_v = VOL_v_lus
        VOL_w = VOL_w_lus
        VOL_z = VOL_z_lus
        THINCOEF = THINCOEF_cyp
        MORT_k = MORT_k_lus_NI
        MORT_m = MORT_m_lus_NI
        MORT_n = MORT_n_lus_NI
        Den_a = DEN_a_cyp
        Den_b = DEN_b_cyp
    ElseIf Species = "Cupressus lusitanica (S.I.)" Then
        MTH_model = MTH_model_lus
        MTH_form = MTH_form_lus
        MTH_a = MTH_a_lus
        MTH_b = MTH_b_lus
        MTH_c = MTH_c_lus
        DBH_model = DBH_model_lus
        DBH_form = DBH_form_lus
        DBH_a = DBH_a_lus
        DBH_b = DBH_b_lus
        DBH_c = DBH_c_lus
        DBH_d = DBH_d_lus
        DBH_f = DBH_f_lus
        DBH_g = DBH_g_lus
        DBH_h = DBH_h_lus
        DBH_k = DBH_k_lus
        MTH_MnHt_a = MTH_MnHt_a_lus
        MTH_MnHt_b = MTH_MnHt_b_lus
        VOL_type = VOL_type_lus
        VOL_u = VOL_u_lus
        VOL_v = VOL_v_lus
        VOL_w = VOL_w_lus
        VOL_z = VOL_z_lus
        THINCOEF = THINCOEF_cyp
        MORT_k = MORT_k_lus_SI
        MORT_m = MORT_m_lus_SI
        MORT_n = MORT_n_lus_SI
        Den_a = DEN_a_cyp
        Den_b = DEN_b_cyp
    ElseIf Species = "Blackwood" Then
        MTH_model = MTH_model_bla
        MTH_form = MTH_form_bla
        MTH_a = MTH_a_bla
        MTH_b = MTH_b_bla
        MTH_c = MTH_c_bla
        DBH_model = DBH_model_bla
        DBH_form = DBH_form_bla
        DBH_a = DBH_a_bla
        DBH_b = DBH_b_bla
        DBH_c = DBH_c_bla
        DBH_d = DBH_d_bla
        DBH_f = DBH_f_bla
        DBH_g = DBH_g_bla
        DBH_h = DBH_h_bla
        DBH_k = DBH_k_bla
        MTH_MnHt_a = MTH_MnHt_a_bla
        MTH_MnHt_b = MTH_MnHt_b_bla
        VOL_type = VOL_type_bla
        VOL_u = VOL_u_bla
        VOL_v = VOL_v_bla
        VOL_w = VOL_w_bla
        VOL_z = VOL_z_bla
        THINCOEF = THINCOEF_bla
        MORT_k = MORT_k_bla
        MORT_m = MORT_m_bla
        MORT_n = MORT_n_bla
        Den_a = DEN_a_bla
        Den_b = DEN_b_bla
    ElseIf Species = "Eucalyptus regnans" Then
        MTH_model = MTH_model_reg
        MTH_form = MTH_form_reg
        MTH_a = MTH_a_reg
        MTH_b = MTH_b_reg
        MTH_c = MTH_c_reg
        DBH_model = DBH_model_reg
        DBH_form = DBH_form_reg
        DBH_a = DBH_a_reg
        DBH_b = DBH_b_reg
        DBH_c = DBH_c_reg
        DBH_d = DBH_d_reg
        DBH_f = DBH_f_reg
        DBH_g = DBH_g_reg
        DBH_h = DBH_h_reg
        DBH_k = DBH_k_reg
        MTH_MnHt_a = MTH_MnHt_a_reg
        MTH_MnHt_b = MTH_MnHt_b_reg
        VOL_type = VOL_type_reg
        VOL_u = VOL_u_reg
        VOL_v = VOL_v_reg
        VOL_w = VOL_w_reg
        VOL_z = VOL_z_reg
        THINCOEF = THINCOEF_euc
        MORT_k = MORT_k_reg
        MORT_m = MORT_m_reg
        MORT_n = MORT_n_reg
        Den_a = DEN_a_reg
        Den_b = DEN_b_reg
    ElseIf Species = "Eucalyptus fastigata" Then
        MTH_model = MTH_model_fas
        MTH_form = MTH_form_fas
        MTH_a = MTH_a_fas
        MTH_b = MTH_b_fas
        MTH_c = MTH_c_fas
        DBH_model = DBH_model_fas
        DBH_form = DBH_form_fas
        DBH_a = DBH_a_fas
        DBH_b = DBH_b_fas
        DBH_c = DBH_c_fas
        DBH_d = DBH_d_fas
        DBH_f = DBH_f_fas
        DBH_g = DBH_g_fas
        DBH_h = DBH_h_fas
        DBH_k = DBH_k_fas
        MTH_MnHt_a = MTH_MnHt_a_fas
        MTH_MnHt_b = MTH_MnHt_b_fas
        VOL_type = VOL_type_fas
        VOL_u = VOL_u_fas
        VOL_v = VOL_v_fas
        VOL_w = VOL_w_fas
        VOL_z = VOL_z_fas
        THINCOEF = THINCOEF_euc
        MORT_k = MORT_k_fas
        MORT_m = MORT_m_fas
        MORT_n = MORT_n_fas
        Den_a = DEN_a_fas
        Den_b = DEN_b_fas
    ElseIf Species = "Eucalyptus nitens (N.I.)" Then
        MTH_model = MTH_model_nit
        MTH_form = MTH_form_nit
        MTH_a = MTH_a_nit
        MTH_b = MTH_b_nit
        MTH_c = MTH_c_nit
        DBH_model = DBH_model_nit
        DBH_form = DBH_form_nit
        DBH_a = DBH_a_nit
        DBH_b = DBH_b_nit
        DBH_c = DBH_c_nit
        DBH_d = DBH_d_nit
        DBH_f = DBH_f_nit
        DBH_g = DBH_g_nit
        DBH_h = DBH_h_nit
        DBH_k = DBH_k_nit
        MTH_MnHt_a = MTH_MnHt_a_nit
        MTH_MnHt_b = MTH_MnHt_b_nit
        VOL_type = VOL_type_nit
        VOL_u = VOL_u_nit
        VOL_v = VOL_v_nit
        VOL_w = VOL_w_nit
        VOL_z = VOL_z_nit
        THINCOEF = THINCOEF_euc
        MORT_k = MORT_k_nit_NI
        MORT_m = MORT_m_nit_NI
        MORT_n = MORT_n_nit_NI
        Den_a = DEN_a_nit
        Den_b = DEN_b_nit
    ElseIf Species = "Eucalyptus nitens (S.I.)" Then
        MTH_model = MTH_model_nit
        MTH_form = MTH_form_nit
        MTH_a = MTH_a_nit
        MTH_b = MTH_b_nit
        MTH_c = MTH_c_nit
        DBH_model = DBH_model_nit
        DBH_form = DBH_form_nit
        DBH_a = DBH_a_nit
        DBH_b = DBH_b_nit
        DBH_c = DBH_c_nit
        DBH_d = DBH_d_nit
        DBH_f = DBH_f_nit
        DBH_g = DBH_g_nit
        DBH_h = DBH_h_nit
        DBH_k = DBH_k_nit
        MTH_MnHt_a = MTH_MnHt_a_nit
        MTH_MnHt_b = MTH_MnHt_b_nit
        VOL_type = VOL_type_nit
        VOL_u = VOL_u_nit
        VOL_v = VOL_v_nit
        VOL_w = VOL_w_nit
        VOL_z = VOL_z_nit
        THINCOEF = THINCOEF_euc
        MORT_k = MORT_k_nit_SI
        MORT_m = MORT_m_nit_SI
        MORT_n = MORT_n_nit_SI
        Den_a = DEN_a_nit
        Den_b = DEN_b_nit
    ElseIf Species = "Eucalyptus delegatensis" Then
        MTH_model = MTH_model_del
        MTH_form = MTH_form_del
        MTH_a = MTH_a_del
        MTH_b = MTH_b_del
        MTH_c = MTH_c_del
        DBH_model = DBH_model_del
        DBH_form = DBH_form_del
        DBH_a = DBH_a_del
        DBH_b = DBH_b_del
        DBH_c = DBH_c_del
        DBH_d = DBH_d_del
        DBH_f = DBH_f_del
        DBH_g = DBH_g_del
        DBH_h = DBH_h_del
        DBH_k = DBH_k_del
        MTH_MnHt_a = MTH_MnHt_a_del
        MTH_MnHt_b = MTH_MnHt_b_del
        VOL_type = VOL_type_del
        VOL_u = VOL_u_del
        VOL_v = VOL_v_del
        VOL_w = VOL_w_del
        VOL_z = VOL_z_del
        THINCOEF = THINCOEF_euc
        MORT_k = MORT_k_del
        MORT_m = MORT_m_del
        MORT_n = MORT_n_del
        Den_a = DEN_a_del
        Den_b = DEN_b_del
    ElseIf Species = "Eucalyptus saligna" Then
        MTH_model = MTH_model_sal
        MTH_form = MTH_form_sal
        MTH_a = MTH_a_sal
        MTH_b = MTH_b_sal
        MTH_c = MTH_c_sal
        DBH_model = DBH_model_sal
        DBH_form = DBH_form_sal
        DBH_a = DBH_a_sal
        DBH_b = DBH_b_sal
        DBH_c = DBH_c_sal
        DBH_d = DBH_d_sal
        DBH_f = DBH_f_sal
        DBH_g = DBH_g_sal
        DBH_h = DBH_h_sal
        DBH_k = DBH_k_sal
        MTH_MnHt_a = MTH_MnHt_a_sal
        MTH_MnHt_b = MTH_MnHt_b_sal
        VOL_type = VOL_type_sal
        VOL_u = VOL_u_sal
        VOL_v = VOL_v_sal
        VOL_w = VOL_w_sal
        VOL_z = VOL_z_sal
        THINCOEF = THINCOEF_euc
        MORT_k = MORT_k_sal
        MORT_m = MORT_m_sal
        MORT_n = MORT_n_sal
        Den_a = DEN_a_sal
        Den_b = DEN_b_sal
    ElseIf Species = "Radiata pine" Then
        MTH_MnHt_a = MTH_MnHt_a_rad
        MTH_MnHt_b = MTH_MnHt_b_rad
        VOL_type = VOL_type_rad
        VOL_u = VOL_u_rad
        VOL_v = VOL_v_rad
        VOL_w = VOL_w_rad
        VOL_z = VOL_z_rad
        THINCOEF = THINCOEF_rad
    ElseIf Species = "Douglas-fir" Then
        MTH_MnHt_a = MTH_MnHt_a_dfr
        MTH_MnHt_b = MTH_MnHt_b_dfr
        VOL_type = VOL_type_dfr
        VOL_u = VOL_u_dfr
        VOL_v = VOL_v_dfr
        VOL_w = VOL_w_dfr
        VOL_z = VOL_z_dfr
        THINCOEF = THINCOEF_dfr
    End If
    
    rotlength = Cells(6, 5)
    If rotlength > 200 Then rotlength = 200 'Maximum allowed rotation length is 200 years
    Stock_hist_T(0) = 0
    Stock_hist_N(0) = Cells(5, 5)
    For i = 1 To 4
        Stock_hist_Type(i) = Cells(9, 4 + i)    ' Thining type: 1 = waste thin; 2 = production thin
        Stock_hist_T(i) = Int(Cells(11, 4 + i)) ' Ages of thinning must be in integer years
        thin_age(i) = Stock_hist_T(i)
        Stock_hist_N(i) = Cells(12, 4 + i)
        Stock_hist_thin_coeff(i) = Cells(13, 4 + i)
        If Cells(13, 4 + i) = -999 Then Stock_hist_thin_coeff(i) = THINCOEF
        If Stock_hist_T(i) <> 0 Then Nthins = i
    Next i
    Stock_hist_Type(5) = 1
    Stock_hist_T(5) = 0
    Stock_hist_N(5) = 0
    Stock_hist_thin_coeff(5) = 1
    
    mode = Cells(16, 4) 'Mode: 1 = Use specified indices, 2 = Calibrate using stand metrics, 3 = Calibrate using tree list
    DiaDist = Cells(16, 5) 'Diamater distribution method: 1 = Weibull, 2 = Derive from tree list
    If Cells(26, 5) <> -999 Then MORT_k = Cells(26, 5) / 100
    weibull_CV = Cells(27, 5)
    If Cells(27, 5) = -999 Then weibull_CV = 0.27 ' Default DBH CV
    weibull_b = 1.010369 * weibull_CV ^ (-1.078517) ' Approximate Weibull b parameter from CV
    
    ' Taper function coefficients
    alpha0 = alpha0_458
    alpha1 = alpha1_458
    alpha2 = alpha2_458
    beta1 = beta1_458
    beta2 = beta2_458
    beta3 = beta3_458
    beta4 = beta4_458
    beta5 = beta5_458
    
    ' Other parameters
    log_length = Cells(28, 5)
    If Cells(28, 5) = -999 Then log_length = 6
    min_SED = Cells(29, 5)
    If Cells(29, 5) = -999 Then min_SED = 150
    break_height = Cells(30, 5) / 100
    If Cells(30, 5) = -999 Then break_height = 0.65
    log_losses = Cells(31, 5)
    If Cells(31, 5) = -999 Then log_losses = 4
    WoodDensity_Adjustment = 1 ' No wood density adjustment
    If Cells(32, 5) <> -999 Then WoodDensity_Adjustment = 1 + Cells(32, 5) / 100
    AGCWD_half_life = Cells(33, 5)
    If Cells(33, 5) = -999 Then AGCWD_half_life = 15 'Natural forest NZ
    BGCWD_half_life = Cells(34, 5)
    If Cells(34, 5) = -999 Then BGCWD_half_life = 15 'Natural forest NZ
    latitude = Cells(35, 5)
    If Cells(35, 5) = -999 Then latitude = 36
    elevation = Cells(36, 5)
    If Cells(36, 5) = -999 Then elevation = 200
    Soil_C = Cells(37, 5)
    If Cells(37, 5) = -999 Then Soil_C = 5.57
    Soil_N = Cells(38, 5)
    If Cells(38, 5) = -999 Then Soil_N = 0.296
    MAT = Cells(39, 5)
    If Cells(39, 5) = -999 Then MAT = 12
    drift = Cells(40, 5)
    If Cells(40, 5) = -999 Then drift = 0
    
    ' Pruning information
    PRUNEHT = 0
    For i = 1 To 4
        prune_age(i) = Cells(9, 12 + i)
        prune_N(i) = Cells(10, 12 + i)
        prune_height(i) = Cells(11, 12 + i)
        If prune_height(i) <> 0 Then PRUNEHT = prune_height(i)
    Next i
        
    ' If species is radiata pine, copy inputs into 300 Index worksheet
    If Species = "Radiata pine" Then
        Worksheets("300 Index").Activate
        Range("C3:C4").ClearContents
        Range("F3:F4").ClearContents
        Range("C7:C11").ClearContents
        Range("C14:C15").ClearContents
        Range("C19").ClearContents
        Range("B20:F36").ClearContents
        Range("B40:E44").ClearContents
        Cells(3, 3) = I300
        Cells(4, 3) = H30
        Cells(3, 6) = latitude
        Cells(4, 6) = elevation
        Cells(75, 4) = Soil_C
        Cells(76, 4) = Soil_N
        Cells(77, 4) = MAT
        Cells(7, 3) = T1
        Cells(8, 3) = N1
        Cells(10, 3) = N1 * PI * (D1 / 200) ^ 2
        Cells(14, 3) = T1
        Cells(15, 3) = H1
        Cells(19, 3) = Stock_hist_N(0)
        For i = 1 To 4
            If Stock_hist_T(i) <> 0 Then
                Cells(19 + i, 2) = Stock_hist_T(i) ' Ages of thinning in integer years
                Cells(19 + i, 4) = Stock_hist_N(i)
                Cells(19 + i, 5) = Stock_hist_thin_coeff(i)
                If Stock_hist_thin_coeff(i) = -999 Then Cells(19 + i, 5) = THINCOEF
            End If
        Next i
        For i = 1 To 4
            If prune_age(i) <> 0 Then
                Cells(39 + i, 2) = prune_age(i)
                Cells(39 + i, 3) = prune_height(i)
                Cells(39 + i, 4) = prune_N(i)
            End If
        Next i
        Cells(47, 3) = rotlength
    End If
    
    ' If species is Douglas-fir, copy inputs into 500 Index worksheet
'    If Species = "Douglas-fir" Then
'        Worksheets("500 Index").Activate
'        Range("B2:B3").ClearContents
'        Range("D2:D4").ClearContents
'        Range("F2:F3").ClearContents
'        Range("C8:G10").ClearContents
'        Range("C14:G16").ClearContents
'        Range("J5:J10").ClearContents
'        Cells(2, 2) = I300
'        Cells(3, 2) = H30
'        Cells(2, 4) = latitude
'        Cells(3, 4) = MAT
'        Cells(4, 4) = Soil_C / (Soil_N - 0.14)
'        Cells(5, 10) = T1
'        Cells(9, 10) = N1
'        Cells(7, 10) = H1
'        Cells(6, 10) = N1 * PI * (D1 / 200) ^ 2
'        Cells(2, 6) = Stock_hist_N(0)
'        For i = 1 To 4
'            If Stock_hist_T(i) <> 0 Then
'                Cells(14, 2 + i) = Stock_hist_T(i) ' Ages of thinning in integer years
'                Cells(15, 2 + i) = Stock_hist_N(i)
'                Cells(16, 2 + i) = Stock_hist_thin_coeff(i)
'                If Stock_hist_thin_coeff(i) = -999 Then Cells(16, 2 + i) = THINCOEF
'            End If
'        Next i
'        For i = 1 To 4
'            If prune_age(i) <> 0 Then
'                Cells(8, 2 + i) = prune_age(i)
'                Cells(9, 2 + i) = prune_height(i)
'                Cells(10, 2 + i) = prune_N(i)
'            End If
'        Next i
'        Cells(3, 6) = rotlength
'    End If
End Sub

Sub Yield_Table_radiata()
    Dim T As Double, thin As Integer, MTDia As Double
    Worksheets("300 Index").Activate
    Call OutputGrowth
    thin = 0
    For T = 0 To rotlength
        N(T) = Cells(5 + T, 8)
        MTH(T) = Cells(5 + T, 10)
        Vol(T) = Cells(5 + T, 12)
        BA(T) = Cells(5 + T, 14)
        DBH(T) = Cells(5 + T, 16)
        WoodDen(T) = Cells(5 + T, 40) * 1000 * WoodDensity_Adjustment
        If Cells(5 + T, 9) <> 0 Then
            thin = thin + 1
            Pre_thin_N(thin) = N(T)
            Pre_thin_dbh(thin) = DBH(T)
            Pre_thin_vol(thin) = Vol(T)
            N(T) = Cells(5 + T, 9)
            Vol(T) = Cells(5 + T, 13)
            BA(T) = Cells(5 + T, 15)
            DBH(T) = Cells(5 + T, 17)
            If Pre_thin_N(thin) - N(T) <= 0 Or (Pre_thin_N(thin) * Pre_thin_dbh(thin) ^ 2 - N(T) * DBH(T) ^ 2) <= 0 Then    ' Calculate dbh of thinned stems checking that estimate is real
                Thin_dbh(thin) = Pre_thin_dbh(thin)
            Else
                Thin_dbh(thin) = Sqr((Pre_thin_N(thin) * Pre_thin_dbh(thin) ^ 2 - N(T) * DBH(T) ^ 2) / (Pre_thin_N(thin) - N(T)))
            End If
            Thin_vol(thin) = Pre_thin_vol(thin) - Vol(T)
            Call Scale_tree_list(Treelist(), nstems, Pre_thin_N(thin), Pre_thin_dbh(thin))
            MTDia = MTD(Treelist(), nstems, Pre_thin_N(thin))
            petB = 1.98 ' Mean petB coefficient in PSPs > 20 years old is 1.98
            petA = (MTH(T) - 1.4) ^ (-1 / 2.5) - petB / MTDia  ' Calculate petA coefficent from MTH
            If petA < 0 Then
                petA = 0
                petB = MnHt_from_MTH(MTH(T), N(T), MTH_MnHt_a, MTH_MnHt_b)
            End If
            Call Predict_height(Treelist(), nstems, N(T), DBH(T), MTH(T), PRUNEHT)
        
            If Cali = False Then Call Felled_stems(Species, Treelist(), nstems, T, Pre_thin_N(thin) - N(T), Thin_dbh(thin), Thin_vol(thin), _
                N(T), DBH(T), MTH(T), PRUNEHT)
            If Cali = False And Stock_hist_Type(thin) = 2 Then
                Call Make_Logs(Treelist(), nstems, T, Pre_thin_N(thin) - N(T), Thin_dbh(thin), Thin_vol(thin), _
                    N(T), DBH(T), MTH(T), PRUNEHT)
                log_volume(T) = Harvest_volume
            End If
        End If
        
    Next T
    
    Call Scale_tree_list(Treelist(), nstems, N(rotlength), DBH(rotlength))
    MTDia = MTD(Treelist(), nstems, N(rotlength))
    petB = 1.98 ' Mean petB coefficient in PSPs > 20 years old is 1.98
    petA = (MTH(rotlength) - 1.4) ^ (-1 / 2.5) - petB / MTDia  ' Calculate petA coefficent from MTH
    If petA < 0 Then
        petA = 0
        petB = MnHt_from_MTH(MTH(rotlength), N(rotlength), MTH_MnHt_a, MTH_MnHt_b)
    End If
    Call Predict_height(Treelist(), nstems, N(rotlength), DBH(rotlength), MTH(rotlength), PRUNEHT)
    Call Felled_stems(Species, Treelist(), nstems, rotlength, N(rotlength), DBH(rotlength), Vol(rotlength), _
        N(rotlength), DBH(rotlength), MTH(rotlength), PRUNEHT)
    Call Make_Logs(Treelist(), nstems, rotlength, N(rotlength), DBH(rotlength), Vol(rotlength), _
        N(rotlength), DBH(rotlength), MTH(rotlength), PRUNEHT)
    log_volume(rotlength) = Harvest_volume
    Call earlyield
End Sub

Sub Yield_Table_dfir()
    Dim T As Double, thin As Integer, MTDia As Double, prevvol As Double, currentvol As Double
    
    Call Dfir_Yield
    WoodDen(0) = WoodDen(0) * WoodDensity_Adjustment
    log_volume(T) = 0
   
    For T = 1 To rotlength
        log_volume(T) = 0
   
        For thin = 1 To 4
            If T >= Stock_hist_T(thin) And T < Stock_hist_T(thin) + 1 Then
   
                Call Scale_tree_list(Treelist(), nstems, Pre_thin_N(thin), Pre_thin_dbh(thin))
                MTDia = MTD(Treelist(), nstems, Pre_thin_N(thin))
                petB = 1.98 ' Mean petB coefficient in PSPs > 20 years old is 1.98
                petA = (MTH(T) - 1.4) ^ (-1 / 2.5) - petB / MTDia  ' Calculate petA coefficent from MTH
                If petA < 0 Then
                    petA = 0
                    petB = MnHt_from_MTH(MTH(T), N(T), MTH_MnHt_a, MTH_MnHt_b)
                End If
                Call Predict_height(Treelist(), nstems, N(T), DBH(T), MTH(T), PRUNEHT)
                N(T) = Stock_hist_N(thin)
                If Pre_thin_N(thin) - N(T) <= 0 Or (Pre_thin_N(thin) * Pre_thin_dbh(thin) ^ 2 - N(T) * DBH(T) ^ 2) <= 0 Then    ' Calculate dbh of thinned stems checking that estimate is real
                    Thin_dbh(thin) = Pre_thin_dbh(thin)
                Else
                    Thin_dbh(thin) = Sqr((Pre_thin_N(thin) * Pre_thin_dbh(thin) ^ 2 - N(T) * DBH(T) ^ 2) / (Pre_thin_N(thin) - N(T)))
                End If

                Thin_vol(thin) = Pre_thin_vol(thin) - Vol(T)
                If Cali = False Then Call Felled_stems(Species, Treelist(), nstems, T, Pre_thin_N(thin) - N(T), Thin_dbh(thin), Thin_vol(thin), _
                    N(T), DBH(T), MTH(T), PRUNEHT)
                If Cali = False And Stock_hist_Type(thin) = 2 Then
                    Call Make_Logs(Treelist(), nstems, T, Pre_thin_N(thin) - N(T), Thin_dbh(thin), Thin_vol(thin), _
                        N(T), DBH(T), MTH(T), PRUNEHT)
                    log_volume(T) = Harvest_volume
                End If
            End If
        Next thin
        
        If DBH(T) <= 0 Then
            DBH(T) = 0
'            BA(T) = 0
        Else
'            BA(T) = N(T) * PI * (DBH(T) / 200) ^ 2
'            Vol(T) = Vol_stand(MTH(T), DBH(T), N(T), VOL_type, VOL_u, VOL_v, VOL_w, VOL_z)
            Call Scale_tree_list(Treelist(), nstems, N(T), DBH(T))
            MTDia = MTD(Treelist(), nstems, N(T))
            petB = 1.98 ' Mean petB coefficient in PSPs > 20 years old is 1.98
            petA = (MTH(T) - 1.4) ^ (-1 / 2.5) - petB / MTDia  ' Calculate petA coefficent from MTH
            If petA < 0 Then
                petA = 0
            End If
            Call Predict_height(Treelist(), nstems, N(T), DBH(T), MTH(T), PRUNEHT)
        End If
        WoodDen(T) = WoodDen(T) * WoodDensity_Adjustment
    Next T
    If Cali = False Then Call Felled_stems(Species, Treelist(), nstems, rotlength, N(rotlength), DBH(rotlength), Vol(rotlength), _
        N(rotlength), DBH(rotlength), MTH(rotlength), PRUNEHT)
    If Cali = False Then Call Make_Logs(Treelist(), nstems, rotlength, N(rotlength), DBH(rotlength), Vol(rotlength), _
                        N(rotlength), DBH(rotlength), MTH(rotlength), PRUNEHT)
    log_volume(rotlength) = Harvest_volume
                
End Sub

Public Sub Yield_Table()
    Dim Total_thin_age_shift As Double, T As Double, thin As Integer, Adj_T As Double, Initial_thin_age_shift As Double, _
        Additional_shift As Double, D300 As Double, MTDia As Double
    
    N(0) = Stock_hist_N(0)
    MTH(0) = 0.3
    DBH(0) = 0
    BA(0) = 0
    Vol(0) = 0
    WoodDen(0) = Den_a * WoodDensity_Adjustment
    log_volume(T) = 0
    Total_thin_age_shift = 0
    
    For T = 1 To rotlength
        Adj_T = Adj_T + 1
        N(T) = N_Mort(N(T - 1), DBH(T - 1), 1, MORT_k, MORT_m, MORT_n)
        MTH(T) = MTH_mod(T, H30, MTH_model, MTH_form, MTH_a, MTH_b, MTH_c)
        log_volume(T) = 0
        Adj_T = T - Total_thin_age_shift
        For thin = 1 To 4
            If T >= Stock_hist_T(thin) And T < Stock_hist_T(thin) + 1 Then
                Pre_thin_N(thin) = N(T)
                Pre_thin_dbh(thin) = DBH_mod(Adj_T, D300_30_est, H30, TBH, N(T), DBH_model, DBH_form, DBH_a, DBH_b, DBH_c, DBH_d, DBH_f, DBH_g, DBH_h, DBH_k, _
                        MTH_model, MTH_form, MTH_a, MTH_b, MTH_c)
                Call Scale_tree_list(Treelist(), nstems, Pre_thin_N(thin), Pre_thin_dbh(thin))
                MTDia = MTD(Treelist(), nstems, Pre_thin_N(thin))
                petB = 1.98 ' Mean petB coefficient in PSPs > 20 years old is 1.98
                petA = (MTH(T) - 1.4) ^ (-1 / 2.5) - petB / MTDia  ' Calculate petA coefficent from MTH
                If petA < 0 Then
                    petA = 0
                    petB = MnHt_from_MTH(MTH(T), N(T), MTH_MnHt_a, MTH_MnHt_b)
                End If
                Call Predict_height(Treelist(), nstems, N(T), DBH(T), MTH(T), PRUNEHT)
                Pre_thin_vol(thin) = Vol_stand(MTH(T), Pre_thin_dbh(thin), Pre_thin_N(thin), VOL_type, VOL_u, VOL_v, VOL_w, VOL_z)
                Initial_thin_age_shift = _
                    Thin_age_shift(Adj_T, N(T), Stock_hist_N(thin), Stock_hist_thin_coeff(thin), TBH, D300_30_est, H30)
                Total_thin_age_shift = Total_thin_age_shift + Initial_thin_age_shift
                Adj_T = T - Total_thin_age_shift
                
                If N(T) > Stock_hist_N(thin) Then N(T) = Stock_hist_N(thin) ' Check that stocking following thinning is less than stocking before thinning
                If Adj_T - TBH <= 0 Then
                    DBH(T) = 0
                Else
                    DBH(T) = DBH_mod(Adj_T, D300_30_est, H30, TBH, N(T), DBH_model, DBH_form, DBH_a, DBH_b, DBH_c, DBH_d, DBH_f, DBH_g, DBH_h, DBH_k, _
                        MTH_model, MTH_form, MTH_a, MTH_b, MTH_c)
                End If
                Vol(T) = Vol_stand(MTH(T), DBH(T), N(T), VOL_type, VOL_u, VOL_v, VOL_w, VOL_z)
                If Pre_thin_N(thin) - N(T) <= 0 Or (Pre_thin_N(thin) * Pre_thin_dbh(thin) ^ 2 - N(T) * DBH(T) ^ 2) <= 0 Then    ' Calculate dbh of thinned stems checking that estimate is real
                    Thin_dbh(thin) = Pre_thin_dbh(thin)
                Else
                    Thin_dbh(thin) = Sqr((Pre_thin_N(thin) * Pre_thin_dbh(thin) ^ 2 - N(T) * DBH(T) ^ 2) / (Pre_thin_N(thin) - N(T)))
                End If
                
                Thin_vol(thin) = Pre_thin_vol(thin) - Vol(T)
                If Cali = False And Not Minimal_run Then Call Felled_stems(Species, Treelist(), nstems, T, Pre_thin_N(thin) - N(T), Thin_dbh(thin), Thin_vol(thin), _
                    N(T), DBH(T), MTH(T), PRUNEHT)
                If Cali = False And Not Minimal_run And Stock_hist_Type(thin) = 2 Then
                    Call Make_Logs(Treelist(), nstems, T, Pre_thin_N(thin) - N(T), Thin_dbh(thin), Thin_vol(thin), _
                        N(T), DBH(T), MTH(T), PRUNEHT)
                    log_volume(T) = Harvest_volume
                End If
                    
                Additional_shift = Initial_thin_age_shift * 0.5
                If Additional_shift > 0.25 Then Additional_shift = 0.25
                If Initial_thin_age_shift < 0 Then Additional_shift = 0
            End If
            If T >= Stock_hist_T(thin) + 1 And T < Stock_hist_T(thin) + 2 Then
                Total_thin_age_shift = Total_thin_age_shift + Additional_shift
            End If
        Next thin
        
        Adj_T = T - Total_thin_age_shift
        If Adj_T - TBH <= 0 Then
            DBH(T) = 0
        Else
            DBH(T) = DBH_mod(Adj_T, D300_30_est, H30, TBH, N(T), DBH_model, DBH_form, DBH_a, DBH_b, DBH_c, DBH_d, DBH_f, DBH_g, DBH_h, DBH_k, _
                MTH_model, MTH_form, MTH_a, MTH_b, MTH_c)
        End If
        
        T_adj(T) = Adj_T
        If DBH(T) <= 0 Then
            DBH(T) = 0
            BA(T) = 0
            Vol(T) = 0
        Else
            BA(T) = N(T) * PI * (DBH(T) / 200) ^ 2
            Vol(T) = Vol_stand(MTH(T), DBH(T), N(T), VOL_type, VOL_u, VOL_v, VOL_w, VOL_z)
            If Not Minimal_run Then
                Call Scale_tree_list(Treelist(), nstems, N(T), DBH(T))
                MTDia = MTD(Treelist(), nstems, N(T))
                petB = 1.98 ' Mean petB coefficient in PSPs > 20 years old is 1.98
                petA = (MTH(T) - 1.4) ^ (-1 / 2.5) - petB / MTDia  ' Calculate petA coefficent from MTH
                If petA < 0 Then
                    petA = 0
                End If
                Call Predict_height(Treelist(), nstems, N(T), DBH(T), MTH(T), PRUNEHT)
            End If
        End If
        
        WoodDen(T) = (Den_a + Den_b * Log(T)) * WoodDensity_Adjustment
    Next T
    If Cali = False And Not Minimal_run Then Call Felled_stems(Species, Treelist(), nstems, rotlength, N(rotlength), DBH(rotlength), Vol(rotlength), _
        N(rotlength), DBH(rotlength), MTH(rotlength), PRUNEHT)
    If Cali = False And Not Minimal_run Then Call Make_Logs(Treelist(), nstems, rotlength, N(rotlength), DBH(rotlength), Vol(rotlength), _
        N(rotlength), DBH(rotlength), MTH(rotlength), PRUNEHT)
    log_volume(rotlength) = Harvest_volume
    Call earlyield
End Sub

Public Sub Calibrate()
    Dim D300_30lo As Double, D300_30up As Double, D300_30mid As Double, Pred_D As Double, i As Integer, j As Integer, _
        flo As Double, fup As Double, fmid As Double, Rotlength_temp As Double
        
    Cali = True
    
    If Species = "Coast redwood" And T2 <> 0 Then MTH_b = MTHmodel_b(T1, H1, T2, H2)
    
    H30 = SI_eqn(T1, H1, MTH_model, MTH_form, MTH_a, MTH_b, MTH_c)
    TBH = AgeBH(30, H30, MTH_model, MTH_form, MTH_a, MTH_b, MTH_c)

    'Make stocking history consistent with calibration measurement
    For i = 1 To Nthins
        If Int(T1) < Stock_hist_T(i) Then
            For j = i To Nthins
                Stock_hist_T(j + 1) = Stock_hist_T(j)
                Stock_hist_N(j + 1) = Stock_hist_N(j)
                Stock_hist_thin_coeff(j + 1) = Stock_hist_thin_coeff(j)
            Next j
            Stock_hist_T(i) = Int(T1)
            Stock_hist_N(i) = N1
            Stock_hist_thin_coeff(i) = 1
            Exit For
        End If
        If Int(T1) = Stock_hist_T(i) Then
            Stock_hist_T(i) = Int(T1)
            Stock_hist_N(i) = N1
            Stock_hist_thin_coeff(i) = 1
            Exit For
        End If
    Next i
    If Int(T1) > Stock_hist_T(Nthins) Then
        Stock_hist_T(Nthins + 1) = Int(T1)
        Stock_hist_N(Nthins + 1) = N1
        Stock_hist_thin_coeff(Nthins + 1) = 1
    End If
    
    'DBH model
    'Estimate D300_30
    D300_30lo = 10
    D300_30up = 120
    If DBH_model = "Korf" And DBH_form = "CA" Then D300_30up = 80
    If DBH_model = "Richards" And DBH_form = "Anamorphic" Then D300_30up = 90
    Rotlength_temp = rotlength  'Temporarily raise rotation length to 200 years allowing for calibration age up to 200 years
    rotlength = 200
    For i = 1 To 16
        D300_30mid = (D300_30lo + D300_30up) / 2
        
        D300_30_est = D300_30lo
        Call Yield_Table
        Pred_D = DBH(Int(T1)) + (DBH(Int(T1 + 1)) - DBH(Int(T1))) * (T1 - Int(T1))
        flo = Pred_D - D1
        
        D300_30_est = D300_30up
        Call Yield_Table
        Pred_D = DBH(Int(T1)) + (DBH(Int(T1 + 1)) - DBH(Int(T1))) * (T1 - Int(T1))
        fup = Pred_D - D1
        
        D300_30_est = D300_30mid
        Call Yield_Table
        Pred_D = DBH(Int(T1)) + (DBH(Int(T1 + 1)) - DBH(Int(T1))) * (T1 - Int(T1))
        fmid = Pred_D - D1
        
        If flo * fmid < 0 Then D300_30up = D300_30mid Else D300_30lo = D300_30mid
    Next i
    
    If Check_errors Then
        If N(Int(T1 - 1)) - N(Int(T1)) > 100 Then _
            MsgBox ("Warning: Stocking of calibration measurement is lower than expected based on stocking after last thinning or at planting - check these and if ncessary adjust them")
        If N(Int(T1 - 1)) - N(Int(T1)) < -100 Then _
            MsgBox ("Warning: Stocking of calibration measurement is higher than expected based on stocking after last thinning or at planting - check these and if ncessary adjust them")
    End If
    
    D300_30_est = D300_30mid
    I300 = I300_from_SI_D300_30(H30, D300_30_est, VOL_type, VOL_u, VOL_v, VOL_w, VOL_z)
    
    Worksheets("Growth Model").Activate
    Cells(3, 5) = I300
    Cells(4, 5) = H30
    rotlength = Rotlength_temp
    Cali = False

End Sub

Sub Calibrate_radiata()
'    Cali = True
    Worksheets("300 Index").Activate
    Call siteIndex
    Call Calc300Index
    Call OutputGrowth
    I300 = Cells(3, 3)
    H30 = Cells(4, 3)
    Worksheets("Growth Model").Activate
    Cells(3, 5) = I300
    Cells(4, 5) = H30
'    Cali = False
End Sub

Sub Calibrate_dfir()
'    Cali = True
    Call CombineSolver
    I300 = I500
    H30 = SI
    Worksheets("Growth Model").Activate
    Cells(3, 5) = I300
    Cells(4, 5) = H30
'    Cali = False
'    Call Dfir_Yield
End Sub

'Estimate b-parameter of Korf MTH model from two MTH measurements using the bisection method
Public Function MTHmodel_b(age1 As Double, MTH1 As Double, Age2 As Double, MTH2 As Double)
    Dim bup As Double, blo As Double, bmid As Double, _
        fup As Double, flo As Double, fmid As Double, i As Integer, ah As Double, _
        pred_MTH2lo As Double, pred_MTH2up As Double, pred_MTH2mid As Double
    blo = 10
    bup = 100
    For i = 1 To 16
        bmid = (blo + bup) / 2
        pred_MTH2lo = 0.3 + Y(Age2, MTH_model, MTH_form, MTH1 - 0.3, age1, MTH_a, blo, MTH_c)
        pred_MTH2up = 0.3 + Y(Age2, MTH_model, MTH_form, MTH1 - 0.3, age1, MTH_a, bup, MTH_c)
        pred_MTH2mid = 0.3 + Y(Age2, MTH_model, MTH_form, MTH1 - 0.3, age1, MTH_a, bmid, MTH_c)
        flo = pred_MTH2lo - MTH2
        fup = pred_MTH2up - MTH2
        fmid = pred_MTH2mid - MTH2
        If fmid * flo < 0 Then bup = bmid Else blo = bmid
    Next i
    MTHmodel_b = bmid
End Function


Function Y(T As Double, model As String, form As String, Y0 As Double, t0 As Double, A As Double, B As Double, C As Double)
    Dim R0 As Double
    If model = "Richards" Then
        If form = "Anamorphic" Then
            Y = Y0 * ((1 - Exp(-B * T)) / (1 - Exp(-B * t0))) ^ (C)
        ElseIf form = "CA" Then
            Y = A * (1 - (1 - (Y0 / A) ^ (1 / C)) ^ (T / t0)) ^ C
        ElseIf form = "GADA" Then
            R0 = (Log(Y0) - Log((1 - Exp(-B * t0)) ^ (C))) / (1 + Log((1 - Exp(-B * t0)) ^ (A)))
            Y = Exp(R0) * (1 - Exp(-B * T)) ^ (C + A * R0)
        Else
            Y = 0
        End If
    ElseIf model = "Korf" Then
        If form = "Anamorphic" Then
            Y = Y0 * Exp(-B * (T) ^ (-C)) / Exp(-B * (t0) ^ (-C))
        ElseIf form = "CA" Then
            Y = A * (Y0 / A) ^ ((t0 / T) ^ (C))
        ElseIf form = "GADA" Then
            R0 = Log(Y0) + Sqr((Log(Y0)) ^ 2 + 4 * B * (t0) ^ (-C))
            Y = Exp(R0 / 2 - 2 * B / (R0 * ((T) ^ (C))))
        Else
            Y = 0
        End If
    ElseIf model = "Hossfeld" Then
        If form = "Anamorphic" Then
            Y = ((T) ^ (C)) / (B + (T) ^ (C) * (1 / Y0 - B / ((t0) ^ (C))))
        ElseIf form = "CA" Then
            Y = ((T) ^ (C)) / (((t0) ^ (C)) / Y0 + ((T) ^ (C) - (t0) ^ (C)) / A)
        ElseIf form = "GADA" Then
            R0 = (Y0 - A + Sqr((Y0 - A) ^ 2 + 4 * Y0 * B * (t0) ^ (-C))) / 2
            Y = (A + R0) / (1 + B * (T) ^ (-C) / R0)
        Else
            Y = 0
        End If
    Else
        Y = 0
    End If
End Function

Function Y0(t0 As Double, model As String, form As String, Y As Double, T As Double, A As Double, B As Double, C As Double)
    Dim R As Double
    If model = "Richards" Then
        If form = "Anamorphic" Then
            Y0 = Y * ((1 - Exp(-B * T)) / (1 - Exp(-B * t0))) ^ (-C)
        ElseIf form = "CA" Then
            Y0 = A * (1 - (1 - (Y / A) ^ (1 / C)) ^ (t0 / T)) ^ C
        ElseIf form = "GADA" Then
            R = (Log(Y) - C * Log(1 - Exp(-B * T))) / (1 + A * Log(1 - Exp(-B * T)))
            Y0 = Exp(R) * (1 - Exp(-B * t0)) ^ (A * R + C)
        Else
            Y0 = 0
        End If
    ElseIf model = "Korf" Then
        If form = "Anamorphic" Then
            Y0 = Y * Exp(-B * (t0) ^ (-C)) / Exp(-B * (T) ^ (-C))
        ElseIf form = "CA" Then
            Y0 = A * (Y / A) ^ ((T / t0) ^ (C))
        ElseIf form = "GADA" Then
            R = Log(Y) + Sqr((Log(Y)) ^ 2 + 4 * B * (T) ^ (-C))
            Y0 = Exp(R / 2 - 2 * B * (t0) ^ (-C) / R)
        Else
            Y0 = 0
        End If
    ElseIf model = "Hossfeld" Then
        If form = "Anamorphic" Then
            Y0 = 1 / (1 / Y + B * (1 / (t0) ^ (C) - 1 / (T) ^ (C)))
        ElseIf form = "CA" Then
            Y0 = ((t0) ^ C) / ((T) ^ C / Y - ((T) ^ C - (t0) ^ C) / A)
        ElseIf form = "GADA" Then
            R = (Y - A + Sqr((Y - A) ^ 2 + 4 * Y * B * (T) ^ (-C))) / 2
            Y0 = R * (A + R) / (R + B * (t0) ^ (-C))
        Else
            Y0 = 0
        End If
    Else
        Y0 = 0
    End If
End Function

Function TP(Y1 As Double, model As String, form As String, Y0 As Double, t0 As Double, A As Double, B As Double, C As Double)
    Dim R0 As Double
    If model = "Richards" Then
        If form = "Anamorphic" Then
            TP = -(1 / B) * Log(1 - (1 - Exp(-B * t0)) * (Y1 / Y0) ^ (1 / C))
        ElseIf form = "CA" Then
            TP = t0 * Log(1 - (Y1 / A) ^ (1 / C)) / Log(1 - (Y0 / A) ^ (1 / C))
        ElseIf form = "GADA" Then
            R0 = (Log(Y0) - Log((1 - Exp(-B * t0)) ^ (C))) / (1 + Log((1 - Exp(-B * t0)) ^ (A)))
            TP = -(1 / B) * Log(1 - Exp((Log(Y1) - R0) / (C + A * R0)))
        Else
            TP = 0
        End If
    ElseIf model = "Korf" Then
        If form = "Anamorphic" Then
            TP = (-1 / B * Log(Y1 / Y0) + (t0) ^ (-C)) ^ (-1 / C)
        ElseIf form = "CA" Then
            TP = t0 * ((Log(Y1 / A)) / (Log(Y0 / A))) ^ (-1 / C)
        ElseIf form = "GADA" Then
            R0 = (Log(Y0) + Sqr((Log(Y0)) ^ 2 + 4 * B * (t0) ^ (-C)))
            TP = (-2 * B / (R0 * Log(Y1) - ((R0) ^ 2) / 2)) ^ (1 / C)
        Else
            TP = 0
        End If
    ElseIf model = "Hossfeld" Then
        If form = "Anamorphic" Then
            TP = (1 / B * (1 / Y1 - 1 / Y0 + B / ((t0) ^ C))) ^ (-1 / C)
        ElseIf form = "CA" Then
            TP = t0 * ((1 / Y0 - 1 / A) / (1 / Y1 - 1 / A)) ^ (1 / C)
        ElseIf form = "GADA" Then
            R0 = (Y0 - A + Sqr((Y0 - A) ^ 2 + 4 * Y0 * B * (t0) ^ (-C))) / 2
            TP = (R0 / B * ((A + R0) / Y1 - 1)) ^ (-1 / C)
        Else
            TP = 0
        End If
    End If
End Function

Public Function CrownHeight(Meanht, N)
    ' Original redwood crown length model
    CrownHeight = pCH + qch * Meanht + rch * Log(N)
    If CrownHeight < 0 Then CrownHeight = 0
End Function

Sub Output_table()
    Dim T As Double, thin As Integer, i As Long, j As Integer
    Worksheets("Growth model").Activate
    Range("J6:AC106").ClearContents
    
    For T = 0 To rotlength
        Cells(6 + T, 10) = T
        Cells(6 + T, 11) = N(T)
        Cells(6 + T, 12) = BA(T)
        Cells(6 + T, 13) = DBH(T)
        Cells(6 + T, 14) = Vol(T)
        Cells(6 + T, 15) = log_volume(T)
        Cells(6 + T, 16) = MTH(T)
        Cells(6 + T, 17) = MnHt_from_MTH(MTH(T), N(T), MTH_MnHt_a, MTH_MnHt_b)
        Cells(6 + T, 18) = WoodDen(T)
        Cells(6 + T, 19) = CrownHeight(MnHt_from_MTH(MTH(T), N(T), MTH_MnHt_a, MTH_MnHt_b), N(T))
        Cells(6 + T, 20) = Worksheets("C Change").Cells(6 + T, 16) * 1 'C 3.667
        Cells(6 + T, 21) = Worksheets("C Change").Cells(6 + T, 17) * 1 'C 3.667
        Cells(6 + T, 22) = Worksheets("C Change").Cells(6 + T, 18) * 1 'C 3.667
        Cells(6 + T, 23) = Worksheets("C Change").Cells(6 + T, 19) * 1 'C 3.667
        Cells(6 + T, 24) = Worksheets("C Change").Cells(6 + T, 20) * 1 'C 3.667
        Cells(6 + T, 25) = Worksheets("C Change").Cells(6 + T, 21) * 1 'C 3.667
        Cells(6 + T, 26) = Worksheets("C Change").Cells(6 + T, 22) * 1 'C 3.667
        Cells(6 + T, 27) = Worksheets("C Change").Cells(6 + T, 23) * 1 'C 3.667
        Cells(6 + T, 28) = Worksheets("C Change").Cells(6 + T, 24) * 1 'C 3.667
        Cells(6 + T, 29) = Worksheets("C Change").Cells(6 + T, 25) * 1 'C 3.667
    Next T
           
    Worksheets("Felled stems").Activate
    Range("A3:F10000").ClearContents
    For i = 1 To stemno - 1
        For j = 1 To 6
            Cells(i + 2, j) = Stemlist(i, j)
        Next j
    Next i
           
    Worksheets("Harvested logs").Activate
    Range("A3:H10000").ClearContents
    For i = 1 To logno - 1
        For j = 1 To 6
            Cells(i + 2, j) = Logs(i, j)
        Next j
'        For j = 9 To 10
'            Cells(i + 2, j - 2) = Logs(i, j)
'        Next j
        For j = 7 To 8
            Cells(i + 2, j) = Logs(i, j)
        Next j
    Next i
    
    Worksheets("Harvest summary").Activate
    Range("A4:M100").ClearContents
    For i = 1 To 5
        If harvest_sum((i - 1) * 4 + 1, 0) <> 0 Then
            Cells((i - 1) * 4 + 4, 1) = harvest_sum((i - 1) * 4 + 1, 0)
            Cells((i - 1) * 4 + 4, 2) = "Butt log"
            Cells((i - 1) * 4 + 5, 2) = "2nd log"
            Cells((i - 1) * 4 + 6, 2) = "Upper logs"
            Cells((i - 1) * 4 + 7, 2) = "Total"
            For j = 1 To 11
                Cells((i - 1) * 4 + 4, j + 2) = harvest_sum((i - 1) * 4 + 1, j)
                Cells((i - 1) * 4 + 5, j + 2) = harvest_sum((i - 1) * 4 + 2, j)
                Cells((i - 1) * 4 + 6, j + 2) = harvest_sum((i - 1) * 4 + 3, j)
                Cells((i - 1) * 4 + 7, j + 2) = harvest_sum((i - 1) * 4 + 4, j)
            Next j
        End If
    Next i
    
    Worksheets("Growth model").Activate
End Sub

Sub Kizha_Han()
    Dim CFraction_wood As Double, CFraction_bark As Double, CFraction_roots As Double, CFraction_needles As Double, _
        CFraction_branches As Double, CFraction_DWL As Double, root_shoot_ratio As Double, i As Long, Age As Double, _
        stocking As Double, Vol As Double, DBH As Double, MTH As Double, dead_vol As Double, post_thin_stocking As Double, _
        post_thin_vol As Double, post_thin_dbh As Double, biomass_stem_wood As Double, biomass_bark As Double, _
        biomass_live_branches As Double, biomass_dead_branches As Double, biomass_foliage As Double, AGB As Double, _
        BGB As Double, AGCWD As Double, BGCWD As Double, pre_thin_biomass_stem_wood As Double, pre_thin_biomass_bark As Double, _
        pre_thin_BGB As Double, tree As Long, bark_wood_ratio As Double, bark_density As Double, harvest_vol As Double, _
        biomass_harvested As Double, AGCWD_rot_1 As Double, BGCWD_rot_1 As Double, wooddensity As Double
        
    Worksheets("C Change").Activate
    CFraction_wood = 0.53   'Jones & OHara 2012
    CFraction_bark = 0.519   'Wilson, Funck & Avery 2010
    CFraction_roots = 0.53  'Same as wood
    CFraction_needles = 0.495   'Van Pelt 2016
    CFraction_branches = 0.495  'Van Pelt 2016
    CFraction_DWL = 0.5 'IPCC default
    bark_wood_ratio = 0.18  'Miles and Smith 2009
    bark_density = 437  'Miles and Smith 2009
    root_shoot_ratio = 0.23 'IPCC default
    For i = 0 To rotlength
        Age = Cells(6 + i, 7)
        stocking = Cells(6 + i, 8)
        Vol = Cells(6 + i, 11)
        DBH = Cells(6 + i, 30)
        MTH = Cells(6 + i, 10)
        wooddensity = Cells(6 + i, 14)
        dead_vol = Cells(6 + i, 13)
        post_thin_stocking = Cells(6 + i, 9)
        post_thin_vol = Cells(6 + i, 12)
        post_thin_dbh = Cells(6 + i, 31)
        harvest_vol = Cells(6 + i, 15)
        If DBH = 0 Then ' Estimate biomass in very young trees
            biomass_stem_wood = wooddensity * Vol / 1000
            biomass_bark = 0
            biomass_live_branches = 0
            biomass_dead_branches = 0
            BGB = root_shoot_ratio * biomass_stem_wood
            AGCWD = 0
            BGCWD = 0
        Else
            Call Scale_tree_list(Treelist(), nstems, stocking, DBH)
            biomass_stem_wood = wooddensity * Vol / 1000
            biomass_bark = Vol * bark_wood_ratio * bark_density / 1000
            biomass_harvested = harvest_vol * wooddensity / 1000 + harvest_vol * bark_wood_ratio * bark_density / 1000
            biomass_live_branches = 0
            biomass_dead_branches = 0
            biomass_foliage = 0
            For tree = 1 To nstems
                biomass_live_branches = biomass_live_branches + (0.01475 * Treelist(tree, 3) ^ 2.0382) / 1000
                biomass_dead_branches = biomass_dead_branches + (0.00038117 * Treelist(tree, 3) ^ 2.3257) / 1000
                biomass_foliage = biomass_foliage + (0.05064 * Treelist(tree, 3) ^ 1.5819) / 1000
            Next tree
            biomass_live_branches = biomass_live_branches / nstems * stocking
            biomass_dead_branches = biomass_dead_branches / nstems * stocking
            biomass_foliage = biomass_foliage / nstems * stocking
            AGB = biomass_stem_wood + biomass_bark + biomass_live_branches + biomass_dead_branches + biomass_foliage
            BGB = AGB * root_shoot_ratio
            AGCWD = AGCWD * 0.5 ^ (1 / AGCWD_half_life)   ' Decay existing AGCWD for 1 year
            BGCWD = BGCWD * 0.5 ^ (1 / BGCWD_half_life)   ' Decay existing BGCWD for 1 year
            AGCWD = AGCWD + (biomass_stem_wood + biomass_bark) * dead_vol / Vol
            BGCWD = BGCWD + BGB * dead_vol / Vol
                ' Add stem and root biomass to CWD from trees that died based on ratio of dead volume to live volume
            If post_thin_stocking <> 0 Then
                pre_thin_biomass_stem_wood = biomass_stem_wood
                pre_thin_biomass_bark = biomass_bark
                pre_thin_BGB = BGB
                stocking = post_thin_stocking
                Vol = post_thin_vol
                DBH = post_thin_dbh
                Call Scale_tree_list(Treelist(), nstems, stocking, DBH)
                biomass_stem_wood = wooddensity * Vol / 1000
                biomass_bark = Vol * bark_wood_ratio * bark_density / 1000
                biomass_live_branches = 0
                biomass_dead_branches = 0
                biomass_foliage = 0
                For tree = 1 To nstems
                    biomass_live_branches = biomass_live_branches + (0.01475 * Treelist(tree, 3) ^ 2.0382) / 1000
                    biomass_dead_branches = biomass_dead_branches + (0.0003817 * Treelist(tree, 3) ^ 2.3257) / 1000
                    biomass_foliage = biomass_foliage + (0.05064 * Treelist(tree, 3) ^ 1.5819) / 1000
                Next tree
                biomass_live_branches = biomass_live_branches / nstems * stocking
                biomass_dead_branches = biomass_dead_branches / nstems * stocking
                biomass_foliage = biomass_foliage / nstems * stocking
                AGB = biomass_stem_wood + biomass_bark + biomass_live_branches + biomass_dead_branches + biomass_foliage
                BGB = AGB * root_shoot_ratio
                AGCWD = AGCWD + pre_thin_biomass_stem_wood - biomass_stem_wood + pre_thin_biomass_bark - biomass_bark - biomass_harvested
                    ' Add to CWD stem, bark & root biomass from thinned trees excluding harvested log stemswood and bark
                BGCWD = BGCWD + pre_thin_BGB - BGB
            End If
        End If
        Cells(6 + i, 20) = 0    ' We do not currently attempt to model fine litter
        Cells(6 + i, 19) = (AGCWD + BGCWD) * CFraction_DWL
        Cells(6 + i, 26) = BGCWD * CFraction_DWL
        Cells(6 + i, 18) = BGB * CFraction_roots
        Cells(6 + i, 17) = biomass_stem_wood * CFraction_wood + biomass_bark * CFraction_bark + _
            biomass_live_branches * CFraction_branches + biomass_dead_branches * CFraction_DWL + _
            biomass_foliage * CFraction_needles
        Cells(6 + i, 16) = Cells(6 + i, 20) + Cells(6 + i, 19) + Cells(6 + i, 18) + Cells(6 + i, 17)
        Cells(6 + i, 33) = biomass_stem_wood * CFraction_wood
        Cells(6 + i, 34) = biomass_bark * CFraction_bark
        Cells(6 + i, 35) = biomass_live_branches * CFraction_branches
        Cells(6 + i, 36) = biomass_dead_branches * CFraction_branches
        Cells(6 + i, 37) = biomass_foliage * CFraction_needles
    Next i
    
    ' 2nd rotation
    
    BGCWD_rot_1 = (BGCWD + BGB) * CFraction_DWL
    AGCWD_rot_1 = (AGCWD + biomass_stem_wood + biomass_bark + biomass_live_branches + biomass_dead_branches - biomass_harvested) * CFraction_DWL
    
    For i = 0 To rotlength
        BGCWD_rot_1 = BGCWD_rot_1 * 0.5 ^ (1 / BGCWD_half_life)   ' Decay BGCWD from rotation 1 for 1 year
        AGCWD_rot_1 = AGCWD_rot_1 * 0.5 ^ (1 / AGCWD_half_life)   ' Decay AGCWD from rotation 1 for 1 year
        Cells(6 + i, 22) = Cells(6 + i, 17)
        Cells(6 + i, 23) = Cells(6 + i, 18)
        Cells(6 + i, 24) = Cells(6 + i, 19) + BGCWD_rot_1 + AGCWD_rot_1
        Cells(6 + i, 25) = Cells(6 + i, 20)
        Cells(6 + i, 21) = Cells(6 + i, 22) + Cells(6 + i, 23) + Cells(6 + i, 24) + Cells(6 + i, 25)
    Next i
    
    Worksheets("Growth model").Activate

End Sub

Sub Initialise_C_Change()
    Dim T As Integer, thin As Integer, lift As Integer
    Worksheets("C Change").Activate
    Range("G6:O200").ClearContents
    Range("P6:AB200").ClearContents
    Range("AD6:AF200").ClearContents
    Range("AG6:AL200").ClearContents
    Range("A11:C15").ClearContents
    Range("A20:D24").ClearContents
    
    If Species = "Radiata pine" Then Cells(3, 3) = "PRAD"
    If Species = "Blackwood" Or Species = "Eucalyptus regnans" Or Species = "Eucalyptus fastigata" Or _
        Species = "Eucalyptus nitens (N.I.)" Or Species = "Eucalyptus nitens (S.I.)" Or Species = "Eucalyptus delegatensis" _
        Or Species = "Eucalyptus saligna" Then Cells(3, 3) = "EUC"
    If Species = "Coast redwood" Then Cells(3, 3) = "RED"
    If Species = "Cupressus macrocarpa (N.I.)" Or Species = "Cupressus macrocarpa (S.I.)" Or Species = "Cupressus lusitanica (N.I.)" _
         Or Species = "Cupressus lusitanica (S.I.)" Then Cells(3, 3) = "CLUS"
    If Species = "Douglas-fir" Then Cells(3, 3) = "PMEN"
    
    ' If species is radiata pine or Douglas-fir, then density is Sheath Density, otherwise it is Who;e Tree Density
    If Species = "Radiata pine" Or Species = "Douglas-fir" Then
        Cells(58, 3) = "S"
    Else
        Cells(58, 3) = "T"
    End If
    
    Cells(4, 3) = Stock_hist_N(0)
    Cells(5, 3) = rotlength
    Cells(6, 3) = rotlength
    
    Cells(6, 7) = 0
    Cells(6, 8) = N(0)
    Cells(6, 10) = MTH(0)
    Cells(6, 11) = Vol(0)
    Cells(6, 14) = WoodDen(0)
    Cells(6, 30) = 0
    
    For thin = 1 To 4
        If Stock_hist_T(thin) <> 0 Then
            If Stock_hist_Type(thin) = 1 Then
                Cells(10 + thin, 1) = "W"
            Else
                Cells(10 + thin, 1) = "P"
            End If
            Cells(10 + thin, 2) = Stock_hist_T(thin)
            Cells(10 + thin, 3) = Stock_hist_N(thin)
        End If
    Next thin
    
    For T = 0 To rotlength
        Cells(6 + T, 7) = T
        Cells(6 + T, 8) = N(T)
        Cells(6 + T, 10) = MTH(T)
        Cells(6 + T, 11) = Vol(T)
        Cells(6 + T, 14) = WoodDen(T)
        Cells(6 + T, 15) = log_volume(T)
        Cells(6 + T, 30) = DBH(T)
    Next T
    
    For thin = 1 To 4
        If Stock_hist_T(thin) <> 0 Then
            Cells(6 + Stock_hist_T(thin), 9) = Cells(6 + Stock_hist_T(thin), 8)
            Cells(6 + Stock_hist_T(thin), 8) = Pre_thin_N(thin)
            Cells(6 + Stock_hist_T(thin), 12) = Cells(6 + Stock_hist_T(thin), 11)
            Cells(6 + Stock_hist_T(thin), 11) = Pre_thin_vol(thin)
            Cells(6 + Stock_hist_T(thin), 31) = Cells(6 + Stock_hist_T(thin), 30)
            Cells(6 + Stock_hist_T(thin), 30) = Pre_thin_dbh(thin)
            Cells(6 + Stock_hist_T(thin), 32) = Thin_dbh(thin)
        End If
    Next thin
    
    For lift = 1 To 4
        If prune_age(lift) <> 0 Then
            Cells(19 + lift, 1) = lift
            Cells(19 + lift, 2) = prune_age(lift)
            Cells(19 + lift, 3) = prune_N(lift)
            Cells(19 + lift, 4) = prune_height(lift)
        End If
    Next lift

End Sub

'Estimate initial thinning age-shift using the bisection method
Public Function Thin_age_shift(Age As Double, N1 As Double, N2 As Double, Thin_coeff As Double, Age_BH As Double, _
        D300_30 As Double, SI30 As Double)
    Dim D300 As Double, DBH_pre_thin As Double, DBH_post_thin As Double, Ageup As Double, Agelo As Double, Agemid As Double, _
        D300_30up As Double, D300_30lo As Double, D300_30mid As Double, D300up As Double, D300lo As Double, D300mid As Double, _
        fup As Double, flo As Double, fmid As Double, i As Integer
    
    DBH_pre_thin = DBH_mod(Age, D300_30, SI30, Age_BH, N1, DBH_model, DBH_form, DBH_a, DBH_b, DBH_c, DBH_d, DBH_f, DBH_g, DBH_h, DBH_k, _
                        MTH_model, MTH_form, MTH_a, MTH_b, MTH_c)
    DBH_post_thin = DBH_pre_thin * ((N1 / N2) ^ ((1 - Thin_coeff) / 2))
    
    Agelo = Age_BH + 1
    Ageup = Age - Age_BH + 20
    For i = 1 To 16
        Agemid = (Agelo + Ageup) / 2
        flo = DBH_mod(Agelo, D300_30, SI30, Age_BH, N2, DBH_model, DBH_form, DBH_a, DBH_b, DBH_c, DBH_d, DBH_f, DBH_g, DBH_h, DBH_k, _
                        MTH_model, MTH_form, MTH_a, MTH_b, MTH_c) - DBH_post_thin
        fup = DBH_mod(Ageup, D300_30, SI30, Age_BH, N2, DBH_model, DBH_form, DBH_a, DBH_b, DBH_c, DBH_d, DBH_f, DBH_g, DBH_h, DBH_k, _
                        MTH_model, MTH_form, MTH_a, MTH_b, MTH_c) - DBH_post_thin
        fmid = DBH_mod(Agemid, D300_30, SI30, Age_BH, N2, DBH_model, DBH_form, DBH_a, DBH_b, DBH_c, DBH_d, DBH_f, DBH_g, DBH_h, DBH_k, _
                        MTH_model, MTH_form, MTH_a, MTH_b, MTH_c) - DBH_post_thin
        If flo * fmid < 0 Then Ageup = Agemid Else Agelo = Agemid
    Next i
    Thin_age_shift = Age - Agemid
End Function

Public Sub Process_tree_list()
    ' Read stem measurements from "Starting tree list" into array TreeList
    Dim tree As Long, Plot_area As Double, Age As Double, stocking As Double, MTDia As Double, MTH As Double
    Worksheets("Starting tree list").Activate
    Plot_area = Cells(3, 2)
    Age = Cells(4, 2)
    nstems = Range(Cells(7, 1), Cells(7, 1).End(xlDown)).Rows.Count
    stocking = nstems / Plot_area
    If nstems > 10000 Then nstems = 1
    If nstems > 0 And nstems < 10000 Then
        For tree = 1 To nstems
            Treelist(tree, 2) = 1 / Plot_area ' Weighting of stem in stems/ha
            Treelist(tree, 3) = Cells(tree + 6, 2) ' DBH cm
            Treelist(tree, 4) = Cells(tree + 6, 3) ' Height m
        Next tree
    End If
    MTDia = MTD(Treelist(), nstems, stocking)
    Call FitPettersonType1(petA, petB, Treelist(), nstems)  ' Fit Petterson Height/DBH curve
    MTH = 1.4 + (petA + petB / MTDia) ^ (-2.5)
    Worksheets("Growth model").Activate
    Cells(25, 5) = Age
    Cells(26, 5) = stocking
    Cells(27, 5) = MTH
    Cells(28, 5) = BAlist(Treelist(), nstems, stocking)  ' BA
    Worksheets("Inputs").Activate
    Cells(22, 6) = 1
    Cells(23, 6) = 1
End Sub

Public Function BAlist(Treelist() As Double, nstems As Long, stocking As Double)
    ' Calculate BA of stems in array TreeList
    Dim stem As Long
    BAlist = 0
    For stem = 1 To nstems
        BAlist = BAlist + Treelist(stem, 3) * Treelist(stem, 3) * PI / 40000
    Next stem
    BAlist = BAlist * stocking / nstems
End Function

Public Function Vollist(Treelist() As Double, nstems As Long, stocking As Double)
    ' Calculate stem volume of stems in array TreeList
    Dim stem As Long
    Vollist = 0
    For stem = 1 To nstems
        Vollist = Vollist + Treelist(stem, 6)
    Next stem
    Vollist = Vollist * stocking / nstems
End Function

Sub Input_tree_list()
    ' Generate unscaled DBHs for tree list - either use a Weibull distribution or read from user-supplied tree list
    Dim tree As Long, proportion As Double
    If DiaDist = 1 Then
        ' Generate 100 stems from Weibull distribution
        nstems = 100
        proportion = 0.005
        For tree = 1 To nstems
            Treelist(tree, 3) = (-Log(1 - proportion)) ^ (1 / weibull_b)
            proportion = proportion + 0.01
        Next tree
    End If
    If DiaDist = 2 Then
        ' Use simple scaling for DBH in treelist based on analysis showing DBH CV remains constant over time
        Worksheets("Starting tree list").Activate
        nstems = Range(Cells(7, 1), Cells(7, 1).End(xlDown)).Rows.Count
        If nstems > 10000 Then nstems = 1
        If nstems > 0 And nstems < 10000 Then
            For tree = 1 To nstems
                Treelist(tree, 3) = Cells(tree + 6, 2) ' DBH cm
            Next tree
        End If
        Worksheets("Growth model").Activate
    End If
End Sub

Sub Scale_tree_list(Treelist() As Double, nstems As Long, stocking As Double, DBH As Double)
    Dim DBH_scaling_factor As Double, MTDia As Double, sum_sqd As Double, tree As Long, petA As Double, petB As Double
    sum_sqd = 0
    For tree = 1 To nstems
        sum_sqd = sum_sqd + Treelist(tree, 3) ^ 2
    Next tree
    DBH_scaling_factor = DBH / Sqr(sum_sqd / nstems)
    For tree = 1 To nstems
        Treelist(tree, 2) = stocking / nstems
        Treelist(tree, 3) = Treelist(tree, 3) * DBH_scaling_factor
    Next tree
End Sub

Sub Predict_height(Treelist() As Double, nstems As Long, StandDensity As Double, MnDBH As Double, MTH As Double, PRUNEHT As Double)
    Dim tree As Long
    For tree = 1 To nstems
        Treelist(tree, 5) = 1.4 + (petA + petB / Treelist(tree, 3)) ^ (-2.5)    ' Predicted heights
        Treelist(tree, 6) = vub(Species, Treelist(tree, 3), Treelist(tree, 5), 0, StandDensity, MnDBH, MTH, PRUNEHT)  ' Predicted stem volumes
    Next tree
End Sub

Public Sub Felled_stems(Species As String, Treelist() As Double, nstems As Long, Age As Double, stocking As Double, DBH As Double, Vol As Double, _
    StandDensity As Double, MnDBH As Double, MTH As Double, PRUNEHT As Double)
    Dim i As Integer, stem_vol As Double
    
    Call Scale_tree_list(Treelist(), nstems, stocking, DBH)
    stem_vol = 0
    For i = 1 To nstems
        stem_vol = stem_vol + vub(Species, Treelist(i, 3), Treelist(i, 5), 0, StandDensity, MnDBH, MTH, PRUNEHT)
    Next i
    stem_vol = stem_vol * stocking / nstems
    
    For i = 1 To nstems
        Stemlist(stemno, 1) = Age
        Stemlist(stemno, 2) = i
        Stemlist(stemno, 3) = stocking / nstems
        Stemlist(stemno, 4) = Treelist(i, 3)
        Stemlist(stemno, 5) = Treelist(i, 5)
        Stemlist(stemno, 6) = vub(Species, Treelist(i, 3), Treelist(i, 5), 0, _
            StandDensity, MnDBH, MTH, PRUNEHT) * Vol / stem_vol
        stemno = stemno + 1
    Next i
End Sub

Public Sub Make_Logs(Treelist() As Double, nstems As Long, Age As Double, stocking As Double, DBH As Double, Vol As Double, _
    StandDensity As Double, MnDBH As Double, MTH As Double, PRUNEHT As Double)
    Dim sed As Double, Volume As Double, i As Integer, j As Integer, stem_vol As Double, _
        Small_end_Height_ratio As Double, Large_end_Height_ratio As Double
    
    Call Scale_tree_list(Treelist(), nstems, stocking, DBH)
    Harvest_volume = 0
    stem_vol = 0
    For i = 1 To nstems
'        stem_vol = stem_vol + redwoodvub(Treelist(i, 3), Treelist(i, 5), 0, alpha0, alpha1, alpha2, beta1, beta2, beta3, beta4, beta5)
        stem_vol = stem_vol + vub(Species, Treelist(i, 3), Treelist(i, 5), 0, StandDensity, MnDBH, MTH, PRUNEHT)
    Next i
    stem_vol = stem_vol * stocking / nstems
    
    For i = 1 To nstems
        For j = 1 To 100
            If 0.3 + log_length * j >= Treelist(i, 5) Then Exit For
            sed = dub(Species, Treelist(i, 3), Treelist(i, 5), 0.3 + log_length * j, StandDensity, MnDBH, MTH, PRUNEHT)
            If sed < min_SED / 10 Then Exit For
            If 0.3 + log_length * j >= break_height * Treelist(i, 5) Then Exit For
            Volume = vol_ub(Species, Treelist(i, 3), Treelist(i, 5), 0.3 + log_length * (j - 1), _
                0.3 + log_length * j, StandDensity, MnDBH, MTH, PRUNEHT) * Vol / stem_vol
            Logs(logno, 1) = Age
            Logs(logno, 2) = stocking / nstems * (100 - log_losses) / 100
            Logs(logno, 3) = i
            Logs(logno, 4) = j
            Logs(logno, 5) = sed * 10
            Logs(logno, 6) = 10 * dub(Species, Treelist(i, 3), Treelist(i, 5), 0.3 + log_length * (j - 1), _
                StandDensity, MnDBH, MTH, PRUNEHT)
            Logs(logno, 7) = log_length
            Logs(logno, 8) = Volume
            Harvest_volume = Harvest_volume + Volume * Logs(logno, 2)
'            Small_end_Height_ratio = (Treelist(i, 5) - (0.3 + log_length * j)) / Treelist(i, 5)
'            Large_end_Height_ratio = (Treelist(i, 5) - (0.3 + log_length * (j - 1))) / Treelist(i, 5)
'            Logs(logno, 9) = Logs(logno, 5) * Sqr(uba1 + uba2 * Small_end_Height_ratio + _
                uba3 * Small_end_Height_ratio ^ 2) ' Under-bark SED
'            Logs(logno, 10) = Logs(logno, 6) * Sqr(uba1 + uba2 * Large_end_Height_ratio + _
                uba3 * Large_end_Height_ratio ^ 2) ' Under-bark SED
            logno = logno + 1
        Next j
    Next i
    
    Call harvest_summary
End Sub

Sub harvest_summary()
    Dim last_log_age As Double, log_age As Double, row As Integer, log_size As Integer, _
        log_height As Integer, i As Integer, j As Integer
    last_log_age = Logs(1, 1)
    row = 1
    For i = 1 To 20
        For j = 0 To 11
            harvest_sum(i, j) = 0
        Next j
    Next i
    For i = 1 To logno - 1
        log_age = Logs(i, 1)
        If log_age <> last_log_age Then
            row = row + 4
        End If
        last_log_age = log_age
        log_size = Int(Logs(i, 5) / 100)
        If log_size > 10 Then log_size = 10
        log_height = Logs(i, 4)
        If log_height > 3 Then log_height = 3
        harvest_sum(row, 0) = log_age
        harvest_sum(row + log_height - 1, log_size) = harvest_sum(row + log_height - 1, log_size) + Logs(i, 8) * Logs(i, 2)
        harvest_sum(row + 3, log_size) = harvest_sum(row + 3, log_size) + Logs(i, 8) * Logs(i, 2)
        harvest_sum(row + log_height - 1, 11) = harvest_sum(row + log_height - 1, 11) + Logs(i, 8) * Logs(i, 2)
        harvest_sum(row + 3, 11) = harvest_sum(row + 3, 11) + Logs(i, 8) * Logs(i, 2)
    Next i
End Sub

Sub FitPettersonType1(petA As Double, petB As Double, Treelist() As Double, nstems As Long)
    ' Fit type 1 Petterson Height / DBH curve to stem measurements in TreeList. Return petA and petB
    ' Code supplied by Andrew Gordon
    Dim Nheights As Long, X As Double, Y As Double, sum_x As Double, sum_y As Double, sum_x2 As Double, _
        sum_y2 As Double, sum_xy As Double, stem As Long
    Nheights = 0
    sum_x = 0
    sum_y = 0
    sum_x2 = 0
    sum_y2 = 0
    sum_xy = 0
    For stem = 1 To nstems
        If Treelist(stem, 4) <> 0 And Treelist(stem, 3) <> 0 Then ' base calculation on trees with measured DBH and height
            Y = Treelist(stem, 3) / (Treelist(stem, 4) - 1.4) ^ (0.4)
            X = Treelist(stem, 3)
            Nheights = Nheights + 1
            sum_x = sum_x + X
            sum_y = sum_y + Y
            sum_x2 = sum_x2 + X * X
            sum_y2 = sum_y2 + Y * Y
            sum_xy = sum_xy + X * Y
        End If
    Next stem
    petA = 0
    petB = 0
    If Nheights > 1 Then
        petA = (sum_xy - sum_x * sum_y / Nheights) / (sum_x2 - sum_x ^ 2 / Nheights)
        petB = sum_y / Nheights - petA * (sum_x / Nheights)
        If petB < 0 Then
            petB = 0
            petA = sum_y / sum_x
        End If
        If petA < 0 Then
            petA = 0
            petB = sum_y / Nheights
        End If
    End If
End Sub

Function MTD(Treelist() As Double, nstems As Long, stocking As Double)
    ' Calculate mean top diameter of stem measuremnts in TreeList
    Const nMTD As Long = 100 ' number of top stems for MTD calculation
    Dim stem As Long, Keys() As String, j As Long, wt As Double, sumWt As Double, sumDBH2Wt As Double
    ReDim Keys(nstems)
    For stem = 1 To nstems
        Keys(stem) = Format(Treelist(stem, 3), "0000.000") + Format(stem, "000000")
    Next stem
    Call Quicksort(Keys, 1, nstems)
    sumWt = 0
    sumDBH2Wt = 0
    For j = nstems To 1 Step -1 ' find weighted quadratic mean of nMTD per ha largest stems
        stem = Val(Right(Keys(j), 6))
        If sumWt + stocking / nstems > nMTD Then
            wt = nMTD - sumWt
        Else
            wt = stocking / nstems
        End If
        sumWt = sumWt + wt
        sumDBH2Wt = sumDBH2Wt + wt * Treelist(stem, 3) * Treelist(stem, 3)
        If sumWt >= nMTD Then Exit For
    Next j
    MTD = Sqr(sumDBH2Wt / sumWt)
End Function

Sub Quicksort(List() As String, min As Long, max As Long)
Dim med_value As String
Dim HI As Long
Dim lo As Long
Dim i As Long

    ' If the list has no more than 1 element, it's sorted.
    If min >= max Then Exit Sub

    ' Pick a dividing item.
    ' i = Int((max - min + 1) * Rnd + min)
    i = Int((max + min) / 2)  ' use middle not random
    med_value = List(i)

    ' Swap it to the front so we can find it easily.
    List(i) = List(min)

    ' Move the items smaller than this into the left
    ' half of the list. Move the others into the right.
    lo = min
    HI = max
    Do
        ' Look down from hi for a value < med_value.
        Do While List(HI) >= med_value
            HI = HI - 1
            If HI <= lo Then Exit Do
        Loop
        If HI <= lo Then
            List(lo) = med_value
            Exit Do
        End If

        ' Swap the lo and hi values.
        List(lo) = List(HI)
        
        ' Look up from lo for a value >= med_value.
        lo = lo + 1
        Do While List(lo) < med_value
            lo = lo + 1
            If lo >= HI Then Exit Do
        Loop
        If lo >= HI Then
            lo = HI
            List(HI) = med_value
            Exit Do
        End If

        ' Swap the lo and hi values.
        List(HI) = List(lo)
    Loop

    ' Sort the two sublists
    Quicksort List(), min, lo - 1
    Quicksort List(), lo + 1, max
End Sub

Public Sub earlyield()
' Correct early volume predictions in yield tables
' It is assumed that early volume growth of individual trees is proportional to Age^2.7 (based on data from VMAN)
' This function is applied when DBH is less than 2 cm
' Age zero volume from planted seedlings supplied by Beets
' Volume weightings for these ages were derived using VMAN
    Dim initialvol As Double, Age As Integer, treevolinc As Double, i As Long, _
        ddbh As Double, prevdbh As Double, j As Long, k As Double, TT As Double
    Worksheets("Growth model").Activate
    initialvol = 0.0000064  'Volume of seedling at planting (m3) - Beets
    ddbh = 0
    For Age = 1 To 15
        prevdbh = ddbh
        ddbh = DBH(Age)
        If ddbh >= 2 And prevdbh < 2 Then
            treevolinc = Vol(Age) / N(Age) - initialvol
            k = treevolinc / (Age ^ 2.7)  'parameter k ensures that volume is correct when T=age
            For j = Age - 1 To 0 Step -1
                TT = j
                If TT < 0 Then TT = 0
                Vol(j) = (initialvol + k * TT ^ 2.7) * N(j)
            Next j
        End If
    Next Age
End Sub

Rem returns under bark diameter at level from taper and bark equation
Public Function old_redwooddob(DBH, HT, level, b1, b2, b3, b4, b5)
    Dim z As Double, beta_c As Double, dob As Double
    z = 1 - level / HT
    If level < 0 Then z = 1
    If z < 0.0001 Then z = 0
    beta_c = (1 - b3 / ((DBH * HT) ^ b4) * ((1 - bh / HT) ^ b5)) / ((1 - bh / HT) ^ (b1 / (HT ^ b2)))
    redwooddob = Sqr(DBH * DBH * (beta_c * z ^ (b1 / (HT ^ b2)) + (b3 / ((DBH * HT) ^ b4)) * (z ^ b5)))
End Function

Rem returns under bark diameter at level from taper and bark equation
Public Function old_redwooddub(DBH, HT, level, a0, a1, a2, b1, b2, b3, b4, b5)
    Dim z As Double, beta_c As Double, dob As Double
    z = 1 - level / HT
    If level < 0 Then z = 1
    If z < 0.0001 Then z = 0
    beta_c = (1 - b3 / ((DBH * HT) ^ b4) * ((1 - bh / HT) ^ b5)) / ((1 - bh / HT) ^ (b1 / (HT ^ b2)))
    dob = Sqr(DBH * DBH * (beta_c * z ^ (b1 / (HT ^ b2)) + (b3 / ((DBH * HT) ^ b4)) * (z ^ b5)))
    redwooddub = dob * Sqr(a0 + a1 * z + a2 * z ^ 2)
End Function

Public Function old_redwoodvob(DBH, HT, level, b1, b2, b3, b4, b5)
    Dim l As Double, hp2 As Double, dhp3 As Double, beta1 As Double, gohp2 As Double, p1 As Double, p2 As Double, _
        p3 As Double, p4 As Double, p5 As Double, p6 As Double
        
    l = HT - level
    If level < 0 Then l = HT
    If l / HT < 0.0001 Then l = 0
      
    hp2 = HT ^ b2
    dhp3 = (DBH * HT) ^ b4
    beta1 = (1 - (b3 / dhp3) * (1 - bh / HT) ^ b5) / ((1 - bh / HT) ^ (b1 / hp2))

    gohp2 = b1 / hp2

    p1 = (beta1 / (HT ^ gohp2 * (gohp2 + 1))) * l ^ (gohp2 + 1)
    p2 = (b3 / (dhp3 * HT ^ b5 * (b5 + 1))) * l ^ (b5 + 1#)

    redwoodvob = (PI / 40000) * DBH * DBH * (p1 + p2)

End Function

Public Function old_redwoodvub(DBH, HT, level, a0, a1, a2, b1, b2, b3, b4, b5)
    Dim l As Double, hp2 As Double, dhp3 As Double, beta1 As Double, gohp2 As Double, p1 As Double, p2 As Double, _
        p3 As Double, p4 As Double, p5 As Double, p6 As Double
        
    If DBH = 0 Or HT = 0 Then
        redwoodvub = 0
    Else
        l = HT - level
        If level < 0 Then l = HT
        If l / HT < 0.0001 Then l = 0
      
        hp2 = HT ^ b2
        dhp3 = (DBH * HT) ^ b4
        beta1 = (1 - (b3 / dhp3) * (1 - bh / HT) ^ b5) / ((1 - bh / HT) ^ (b1 / hp2))

        gohp2 = b1 / hp2

        p1 = (a0 * beta1 / (HT ^ gohp2 * (gohp2 + 1))) * l ^ (gohp2 + 1)
        p2 = (a0 * b3 / (dhp3 * HT ^ b5 * (b5 + 1))) * l ^ (b5 + 1#)

        p3 = (a1 * beta1 / (HT ^ (gohp2 + 1) * (gohp2 + 2))) * l ^ (gohp2 + 2)
        p4 = (a1 * b3 / (dhp3 * HT ^ (b5 + 1) * (b5 + 2))) * l ^ (b5 + 2)

        p5 = (a2 * beta1 / (HT ^ (gohp2 + 2) * (gohp2 + 3))) * l ^ (gohp2 + 3)
        p6 = (a2 * b3 / (dhp3 * HT ^ (b5 + 2) * (b5 + 3))) * l ^ (b5 + 3)
    
        redwoodvub = (PI / 40000) * DBH * DBH * (p1 + p2 + p3 + p4 + p5 + p6)
    End If
End Function

' Return volume UB between two levels (m above ground)
Public Function old_vol_ub(DBH, HT, lower, upper, a0, a1, a2, b1, b2, b3, b4, b5)
    vol_ub = redwoodvub(DBH, HT, lower, a0, a1, a2, b1, b2, b3, b4, b5) - redwoodvub(DBH, HT, upper, a0, a1, a2, b1, b2, b3, b4, b5)
End Function


' Taper function T136, calculates diameter under bark at a given height for Douglas-fir
Public Function old_Dfir_CalcDub(V As Double, HT As Double, H As Double)
    ' v is total stem volume (m3), HT is total stem height (m), and H is the height up the stem
    Dim D2 As Double
    Const a1 As Double = 0.319071
    Const a2 As Double = 0
    Const a3 As Double = 23.9972
    Const a4 As Double = -47.47884
    Const a5 As Double = 26.02156
    D2 = 40000 * V / (PI * HT) * (a1 * ((HT - H) / HT) + a2 * ((HT - H) / HT) ^ 2 + a3 * ((HT - H) / HT) ^ 3 + a4 * ((HT - H) / HT) ^ 4 + a5 * ((HT - H) / HT) ^ 5)
    Dfir_CalcDub = (D2) ^ 0.5
End Function

' Partial volume function T136, calculates volume under bark below a given height for Douglas-fir
Public Function old_Dfir_PartTreeVol(V As Double, HT As Double, H As Double)
    ' v is total stem volume (m3), HT is total stem height (m), and H is the height up the stem
    Dim c1 As Double, c2 As Double, c3 As Double, c4 As Double, c5 As EffectParameters
    Const a1 As Double = 0.319071
    Const a2 As Double = 0
    Const a3 As Double = 23.9972
    Const a4 As Double = -47.47884
    Const a5 As Double = 26.02156
    c1 = a1 / 2
    c2 = a2 / 3
    c3 = a3 / 4
    c4 = a4 / 5
    c5 = a5 / 6
    Dfir_PartTreeVol = V - V * (c1 * ((HT - H) / HT) ^ 2 + c2 * ((HT - H) / HT) ^ 3 + c3 * ((HT - H) / HT) ^ 4 + c4 * ((HT - H) / HT) ^ 5 + c5 * ((HT - H) / HT) ^ 6)
End Function


'''''''''''''''''''''''''''''''''''''
'                                   '
' Volume functions (VUB above H)    '
'                                   '
'''''''''''''''''''''''''''''''''''''

Function vub(Species As String, DBH As Double, HT As Double, H As Double, N As Double, MnDBH As Double, MTH As Double, PRUNEHT As Double)
        vub = redwood_vub(DBH, HT, H)
'    If Species = "Coast redwood" Then
'        vub = redwood_vub(DBH, HT, H)
'    ElseIf Species = "Douglas-fir" Then
'        vub = Dfir_Vub(DBH, HT, H)
'    ElseIf H < 15 Then
'        vub = Dfir_Dub(DBH, HT, H)
'    Else
'        vub = Radiata_VUB(DBH, HT, H, N, MnDBH, MTH, PRUNEHT)
'    End If
End Function

Public Function Dfir_Vub(DBH As Double, HT As Double, H As Double)
    Dim c1 As Double, c2 As Double, c3 As Double, c4 As Double, c5 As Double, V As Double
    Const a1 As Double = 0.319071, a2 As Double = 0, a3 As Double = 23.9972, a4 As Double = -47.47884, _
        a5 As Double = 26.02156, v1 As Double = 1.8281198, v2 As Double = 1.102592, v3 As Double = -10.19719
    V = DBH ^ v1 * (HT ^ 2 / (HT - 1.4)) ^ v2 * Exp(v3)
    c1 = a1 / 2
    c2 = a2 / 3
    c3 = a3 / 4
    c4 = a4 / 5
    c5 = a5 / 6
    Dfir_Vub = V * (c1 * ((HT - H) / HT) ^ 2 + c2 * ((HT - H) / HT) ^ 3 + c3 * ((HT - H) / HT) ^ 4 + c4 * ((HT - H) / HT) ^ 5 + c5 * ((HT - H) / HT) ^ 6)
End Function

Public Function redwood_vub(DBH As Double, HT As Double, H As Double)
    Dim hp2 As Double, dhp3 As Double, beta1 As Double, gohp2 As Double, p1 As Double, p2 As Double, _
        p3 As Double, p4 As Double, p5 As Double, p6 As Double
    Const a0 As Double = 0.702, a1 As Double = 0.5646, a2 As Double = -0.6188, b1 As Double = 2.6295, _
        b2 As Double = 0.1406, b3 As Double = 0.1455, b4 As Double = -0.1275, b5 As Double = 22.7873
    If DBH = 0 Or HT = 0 Then
        redwood_vub = 0
    Else
        If H > HT Then H = HT
        H = HT - H
        If H / HT < 0.0001 Then H = 0
      
        hp2 = HT ^ b2
        dhp3 = (DBH * HT) ^ b4
        beta1 = (1 - (b3 / dhp3) * (1 - bh / HT) ^ b5) / ((1 - bh / HT) ^ (b1 / hp2))

        gohp2 = b1 / hp2

        p1 = (a0 * beta1 / (HT ^ gohp2 * (gohp2 + 1))) * H ^ (gohp2 + 1)
        p2 = (a0 * b3 / (dhp3 * HT ^ b5 * (b5 + 1))) * H ^ (b5 + 1#)

        p3 = (a1 * beta1 / (HT ^ (gohp2 + 1) * (gohp2 + 2))) * H ^ (gohp2 + 2)
        p4 = (a1 * b3 / (dhp3 * HT ^ (b5 + 1) * (b5 + 2))) * H ^ (b5 + 2)

        p5 = (a2 * beta1 / (HT ^ (gohp2 + 2) * (gohp2 + 3))) * H ^ (gohp2 + 3)
        p6 = (a2 * b3 / (dhp3 * HT ^ (b5 + 2) * (b5 + 3))) * H ^ (b5 + 3)
    
        redwood_vub = (PI / 40000) * DBH * DBH * (p1 + p2 + p3 + p4 + p5 + p6)
    End If
End Function


' Radiata pine
' Partial tree volume function from Gordon and Budianto
Public Function Radiata_VUB(DBH As Double, HT As Double, H As Double, SPH As Double, MnDBH As Double, MTH As Double, PRHT As Double)
    Dim rspace As Double, sd As Double, FQ As Double, D6 As Double, l As Double, z As Double, zb As Double, _
        zu As Double, g1 As Double, g3 As Double, b3 As Double, b1 As Double, b2 As Double, dob As Double, dub As Double, _
        K1 As Double, K2 As Double
    Const a0 As Double = 0.4242, a01 As Double = -0.002822, a10 As Double = 0.6067, a12 As Double = 0.06129, _
        a2 As Double = -0.207, a31 As Double = 0.3208, bf0 As Double = 0.945, bf1 As Double = -0.387, _
        bf2 As Double = 0.000686, bf3 As Double = -0.267, bf4 As Double = 0.00357, b30 As Double = 0.7768, _
        B31 As Double = -0.1347, g10 As Double = 1.018, g11 As Double = 0.2967, g2 As Double = 12.68, _
        g31 As Double = 1.047
    rspace = 100 / ((SPH) ^ 0.5 * MTH)
    sd = MnDBH ^ 2 / (rspace) ^ 0.5
    FQ = bf0 + bf1 * Exp(-bf2 * sd) + bf3 * Exp(-(HT / MTH) ^ 2) + bf4 * PRHT
    D6 = FQ * DBH
    l = HT - H
    z = l / HT
    zb = 1 - 1.4 / HT
    zu = 1 - 6 / HT
    g1 = g10 + g11 * D6 / (HT - 6)
    g3 = g31 * HT * D6 / DBH
    b3 = b30 + B31 * (DBH - D6) / (6 - 1.4)
    b1 = (1 - (zb ^ g2 / zu ^ g2) * (D6 ^ 2 / DBH ^ 2 - b3 * zu ^ g3) - b3 * zb ^ g3) / (zb ^ g1 - (zb ^ g2 * zu ^ g1) / zu ^ g2)
    b2 = (D6 ^ 2 / DBH ^ 2 - b1 * zu ^ g1 - b3 * zu ^ g3) / zu ^ g2
    
    K1 = (a0 + a01 * HT)
    K2 = 1 + 0.5 / Exp(a12 * HT)
    Radiata_VUB = (PI * DBH ^ 2 * HT / 40000) * ((l / HT) ^ g1 * ((b1 * K1 * l) / ((1 + g1) * HT) + _
        (a10 * b1 * (l / HT) ^ (K2)) / (K2 + g1)) + (l / HT) ^ g2 * ((b2 * K1 * l) / _
        ((1 + g2) * HT) + (a10 * b2 * (l / HT) ^ (K2)) / (K2 + g2)) + (l / HT) ^ g3 * _
        ((b3 * K1 * l) / ((1 + g3) * HT) + (a10 * b3 * (l / HT) ^ (K2)) / (K2 + g3)) + _
        (l / HT) ^ (a31 * HT) * ((a2 * b1 * (l / HT) ^ (1 + g1)) / (1 + g1 + a31 * HT) + _
        (a2 * b2 * (l / HT) ^ (1 + g2)) / (1 + g2 + a31 * HT) + (a2 * b3 * (l / HT) ^ (1 + g3)) / (1 + g3 + a31 * HT)))
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                       '
' UB volume between lower and upper levels on stem      '
'                                                       '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function vol_ub(Species As String, DBH As Double, HT As Double, lower As Double, upper As Double, MnDBH As Double, MTH As Double, N As Double, PRHT As Double)
    vol_ub = vub(Species, DBH, HT, lower, N, MnDBH, MTH, PRHT) - vub(Species, DBH, HT, upper, N, MnDBH, MTH, PRHT)
End Function

'''''''''''''''''''''''''''''''''''''
'                                   '
' Taper functions (DUB)             '
'                                   '
'''''''''''''''''''''''''''''''''''''

Function dub(Species As String, DBH As Double, HT As Double, H As Double, MnDBH As Double, MTH As Double, N As Double, PRUNEHT As Double)
        dub = Redwood_dub(DBH, HT, H)
'    If Species = "Coast redwood" Then
'        dub = Redwood_dub(DBH, HT, H)
'    ElseIf Species = "Douglas-fir" Then
'        dub = Dfir_Dub(DBH, HT, H)
'    ElseIf H < 15 Then
'        dub = Dfir_Dub(DBH, HT, H)
'    Else
'        dub = Radiata_Dub(DBH, HT, H, N, MnDBH, MTH, PRUNEHT)
'    End If
End Function

Rem returns under bark diameter at level from taper and bark equation for Redwood
Public Function Redwood_dub(DBH As Double, HT As Double, level As Double)
    Dim z As Double, beta_c As Double, dob As Double
    Const a0 As Double = 0.702, a1 As Double = 0.5646, a2 As Double = -0.6188, b1 As Double = 2.6295, _
        b2 As Double = 0.1406, b3 As Double = 0.1455, b4 As Double = -0.1275, b5 = 22.7873
    z = 1 - level / HT
    If level < 0 Then z = 1
    If z < 0.0001 Then z = 0
    beta_c = (1 - b3 / ((DBH * HT) ^ b4) * ((1 - bh / HT) ^ b5)) / ((1 - bh / HT) ^ (b1 / (HT ^ b2)))
    dob = Sqr(DBH * DBH * (beta_c * z ^ (b1 / (HT ^ b2)) + (b3 / ((DBH * HT) ^ b4)) * (z ^ b5)))
    Redwood_dub = dob * Sqr(a0 + a1 * z + a2 * z ^ 2)
End Function

'Three point taper function of Gordon and Budianto for radiata
Function Radiata_Dub(DBH, HT, H, SPH, MnDBH, MTH, PRHT)
    Dim rspace As Double, sd As Double, FQ As Double, D6 As Double, l As Double, z As Double, zb As Double, _
        zu As Double, g1 As Double, g3 As Double, b3 As Double, b1 As Double, b2 As Double, dob As Double, dub As Double
    Const a0 As Double = 0.4242, a01 = -0.002822, a10 As Double = 0.6067, a12 As Double = 0.06129, _
        a2 As Double = -0.207, a31 As Double = 0.3208, bf0 As Double = 0.945, bf1 As Double = -0.387, _
        bf2 As Double = 0.000686, bf3 As Double = -0.267, bf4 As Double = 0.00357, b30 As Double = 0.7768, _
        B31 As Double = -0.1347, g10 As Double = 1.018, g11 As Double = 0.2967, g2 As Double = 12.68, _
        g31 As Double = 1.047
    rspace = 100 / ((SPH) ^ 0.5 * MTH)
    sd = MnDBH ^ 2 / (rspace) ^ 0.5
    FQ = bf0 + bf1 * Exp(-bf2 * sd) + bf3 * Exp(-(HT / MTH) ^ 2) + bf4 * PRHT
    D6 = FQ * DBH
    l = HT - H
    z = l / HT
    zb = 1 - 1.4 / HT
    zu = 1 - 6 / HT
    g1 = g10 + g11 * D6 / (HT - 6)
    g3 = g31 * HT * D6 / DBH
    b3 = b30 + B31 * (DBH - D6) / (6 - 1.4)
    b1 = (1 - (zb ^ g2 / zu ^ g2) * (D6 ^ 2 / DBH ^ 2 - b3 * zu ^ g3) - b3 * zb ^ g3) / (zb ^ g1 - (zb ^ g2 * zu ^ g1) / zu ^ g2)
    b2 = (D6 ^ 2 / DBH ^ 2 - b1 * zu ^ g1 - b3 * zu ^ g3) / zu ^ g2
    If (b2 > -2) Then
        dob = (DBH ^ 2 * (b1 * z ^ g1 + b2 * z ^ g2 + b3 * z ^ g3)) ^ 0.5
    Else
        dob = 0
    End If
    dub = dob * (a0 + a01 * HT + a10 * z ^ (Exp(-a12 * HT) / 2) + a2 * z ^ (a31 * HT)) ^ 0.5
    Radiata_Dub = dub
End Function

'Taper function T136, calculates diameter under bark at a given height
Public Function Dfir_Dub(DBH As Double, HT As Double, H As Double)
    Dim D2 As Double, V As Double
    Const a1 As Double = 0.319071, a2 As Double = 0, a3 As Double = 23.9972, a4 As Double = -47.47884, _
        a5 As Double = 26.02156, v1 As Double = 1.8281198, v2 As Double = 1.102592, v3 As Double = -10.19719
    V = DBH ^ v1 * (HT ^ 2 / (HT - 1.4)) ^ v2 * Exp(v3)
    D2 = 40000 * V / (PI * HT) * (a1 * ((HT - H) / HT) + a2 * ((HT - H) / HT) ^ 2 + a3 * ((HT - H) / HT) ^ 3 + a4 * ((HT - H) / HT) ^ 4 + a5 * ((HT - H) / HT) ^ 5)
    Dfir_Dub = (D2) ^ 0.5
End Function

''''''''''''''''''''''''''''''''''''
'                                  '
' Batch mode                       '
'                                  '
''''''''''''''''''''''''''''''''''''

Sub Batch_estimate_indices()
'This subroutine predicts Site Index and 300 Index for PSP data in batch mode
     
    On Error GoTo eh
     
    Dim sp As String, isl As String, thin As Integer, row As Integer, startrow As Integer, endrow As Integer
    Application.ScreenUpdating = False
    
    Worksheets("Batch index estimates").Activate
    startrow = 3
    endrow = Range(Cells(3, 1), Cells(3, 1).End(xlDown)).Rows.Count + 2

    Check_errors = False
    Minimal_run = True
    
    For row = startrow To endrow
        
        Worksheets("Growth model").Activate
        Range(Cells(11, 5), Cells(13, 8)).ClearContents
        
        'Display progress on status bar
        Application.ScreenUpdating = True
        Application.DisplayStatusBar = True ' makes sure that the statusbar is visible
        Application.StatusBar = "Processing row" + Str(row) + " out of" + Str(endrow)
        Application.ScreenUpdating = False
        
        'Read next record
        sp = Worksheets("Batch index estimates").Cells(row, 3)
        isl = Worksheets("Batch index estimates").Cells(row, 2)
        If sp = "AAMEL" Then Worksheets("Growth model").Cells(2, 5) = "Blackwood"
        If sp = "EUREG" Then Worksheets("Growth model").Cells(2, 5) = "Eucalyptus regnans"
        If sp = "EUFAS" Then Worksheets("Growth model").Cells(2, 5) = "Eucalyptus fastigata"
        If sp = "EUNIT" And isl = "NI" Then Worksheets("Growth model").Cells(2, 5) = "Eucalyptus nitens (N.I.)"
        If sp = "EUNIT" And isl = "SI" Then Worksheets("Growth model").Cells(2, 5) = "Eucalyptus nitens (S.I.)"
        If sp = "EUDEL" Then Worksheets("Growth model").Cells(2, 5) = "Eucalyptus delegatensis"
        If sp = "EUSAL" Then Worksheets("Growth model").Cells(2, 5) = "Eucalyptus saligna"

        Worksheets("Growth model").Cells(5, 5) = Worksheets("Batch index estimates").Cells(row, 8)
        If Worksheets("Growth model").Cells(5, 5) > 10000 Then Worksheets("Growth model").Cells(5, 5) = 10000   ' Restrict iniital stocking to max of 10000 stems/ha
        Worksheets("Growth model").Cells(25, 5) = Worksheets("Batch index estimates").Cells(row, 4)
        Worksheets("Growth model").Cells(26, 5) = Worksheets("Batch index estimates").Cells(row, 5)
        Worksheets("Growth model").Cells(27, 5) = Worksheets("Batch index estimates").Cells(row, 7)
        Worksheets("Growth model").Cells(28, 5) = Worksheets("Batch index estimates").Cells(row, 6)
        
        For thin = 1 To 4
            If Worksheets("Batch index estimates").Cells(row, 6 + thin * 4) <> 0 And _
                Worksheets("Batch index estimates").Cells(row, 6 + thin * 4) < Worksheets("Batch index estimates").Cells(row, 4) Then
                Worksheets("Growth model").Cells(11, 4 + thin) = Worksheets("Batch index estimates").Cells(row, 6 + thin * 4)
                Worksheets("Growth model").Cells(12, 4 + thin) = Worksheets("Batch index estimates").Cells(row, 8 + thin * 4)
            End If
        Next thin
        
        Call run_model
        Application.ScreenUpdating = True
        Worksheets("Batch index estimates").Cells(row, 25) = Worksheets("Growth model").Cells(4, 5)
        Worksheets("Batch index estimates").Cells(row, 26) = Worksheets("Growth model").Cells(3, 5)
eh:
    Next row
    
    Application.ScreenUpdating = True
    Application.StatusBar = False   'gives control of the statusbar back to the programme

End Sub

Sub Batch_carbon()
'Estimate carbon at ages 10, 20, 30, 40, 50 and 60 years in batch mode
    Dim Nruns As Long, i As Long, Age As Long
    Worksheets("Batch runs").Activate
    Range("E3:AT1000").ClearContents
    Nruns = Range(Cells(3, 1), Cells(3, 1).End(xlDown)).Rows.Count
    For i = 1 To Nruns
        Worksheets("Growth model").Cells(3, 4) = Worksheets("Batch runs").Cells(i + 2, 1)
        Worksheets("Growth model").Cells(4, 4) = Worksheets("Batch runs").Cells(i + 2, 2)
        Worksheets("Growth model").Cells(5, 4) = Worksheets("Batch runs").Cells(i + 2, 3)
        Worksheets("Growth model").Cells(12, 4) = Worksheets("Batch runs").Cells(i + 2, 4)
        Worksheets("Growth model").Cells(13, 4) = Worksheets("Batch runs").Cells(i + 2, 5)
        Call run_mod_batch
        For Age = 10 To 60 Step 10
            Worksheets("Batch runs").Cells(i + 2, 6 + (Age - 10) / 10 * 6) = Worksheets("Growth model").Cells(6 + Age, 9)
            Worksheets("Batch runs").Cells(i + 2, 7 + (Age - 10) / 10 * 6) = Worksheets("C Change").Cells(6 + Age, 16) * 3.67
            Worksheets("Batch runs").Cells(i + 2, 8 + (Age - 10) / 10 * 6) = Worksheets("C Change").Cells(6 + Age, 17) * 3.67
            Worksheets("Batch runs").Cells(i + 2, 9 + (Age - 10) / 10 * 6) = Worksheets("C Change").Cells(6 + Age, 18) * 3.67
            Worksheets("Batch runs").Cells(i + 2, 10 + (Age - 10) / 10 * 6) = Worksheets("C Change").Cells(6 + Age, 19) * 3.67
            Worksheets("Batch runs").Cells(i + 2, 11 + (Age - 10) / 10 * 6) = Worksheets("C Change").Cells(6 + Age, 20) * 3.67
        Next Age
    Next i
End Sub


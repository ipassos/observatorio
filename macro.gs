Array.prototype.findIndex = function(SearchNC){
  
  if(SearchNC == "") return false;
  
   for(var i = 0; i <this.length; i++)
     if(this[i]==SearchNC) return i;
    
   return -1;
  
};

function Verificar(){
  
  var Entrada = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Entrada');
  var Celulas = Entrada.getRange(5,9).getValue();
  
  if(Celulas <= 0){
    return true
  }else{
    return false
  }
    
}

//function UnCheck() {
  //var Entrada = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Entrada');
  //Entrada.getActiveCell();
  //var intervalo = Entrada.getRangeList([]);
  //intervalo.setValue("0");
//}

function LimparEntrada(){
  
  var Entrada = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Entrada');
  
  Entrada.getActiveCell();
  
  Entrada.getRangeList(['B2:B17', 'D2:D6', 'D8', 'D11','D13', 'F8:F17', 'H2:H13', 'H14', 'H16', 'B19:B23', 'B27:B37', 'B39:B41', 'D19:D26', 'D28:D31', 'D33:D39', 'F19:F41', 'H19:H38', 'B43:B47', 'B51:B61', 'B63:B73', 'B75:B78', 'B80:B86', 'D43:D85', 'F43:F47', 'F51:F61', 'F63:F73', 'F75:F78', 'F80:F86', 'H43:H85', 'B90:B94', 'B98:B108', 'B110:B120', 'B122:B125', 'B127:B133', 'D90:D132', 'F90:F94', 'F98:F108', 'F110:F120', 'F122:F125', 'F127:F133', 'H90:H132', 'B137:B141', 'B145:B155', 'B157:B167', 'B169:B172', 'B174:B180', 'D137:D179', 'F137:F141', 'F145:F155', 'F157:F167', 'F169:F172', 'F174:F180', 'H137:H179']).activate();
  Entrada.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  
  Entrada.getRange('B2').activate();
  

}

function Search(){
  
  var Entrada = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Entrada');
  var Pesquisa = Entrada.getRange('B2').getValue(); // celula com a variavel de busca - NC
  var SheetBanco = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Banco');
  SheetBanco.getActiveCell();
  
  var LocalPesquisa = SheetBanco.getRange(2, 2, SheetBanco.getLastRow()).getValues(); //linha/coluna da variavel NC
  var Resultado = LocalPesquisa.Pesquisa(Pesquisa);
  var Linha = Resultado + 2;
  
  if (Resultado != -1){
    
    var NC= SheetBanco.getRange(Linha, 2).getValue();
    var NUM_FATO= SheetBanco.getRange(Linha, 3).getValue();
    var NUM_OCOR= SheetBanco.getRange(Linha, 4).getValue();
    var TIP_FATO= SheetBanco.getRange(Linha, 5).getValue();
    var NUM_ORG_REGIST= SheetBanco.getRange(Linha, 6).getValue();
    var N_ORG_REGIST= SheetBanco.getRange(Linha, 7).getValue();
    var ORIGEM_COM= SheetBanco.getRange(Linha, 8).getValue();
    var DATA_COM= SheetBanco.getRange(Linha, 9).getValue();
    var HORA_COM= SheetBanco.getRange(Linha, 10).getValue();
    var ANO_REGIST= SheetBanco.getRange(Linha, 11).getValue();
    var FLAGRANTE= SheetBanco.getRange(Linha, 12).getValue();
    var TENTATIVA= SheetBanco.getRange(Linha, 13).getValue();
    var COMUM= SheetBanco.getRange(Linha, 14).getValue();
    var COND= SheetBanco.getRange(Linha, 15).getValue();
    var LAUDO= SheetBanco.getRange(Linha, 16).getValue();
    //var nome_var = SheetBanco.getRange(Linha, col).getValue();
    var NUM_AUT= SheetBanco.getRange(Linha, 17).getValue();
    var NUM_VIT= SheetBanco.getRange(Linha, 18).getValue();
    var NUM_VIT_MORTAS= SheetBanco.getRange(Linha, 19).getValue();
    var NUMTEST= SheetBanco.getRange(Linha, 20).getValue();
    var TIPO_TEST= SheetBanco.getRange(Linha, 21).getValue();
    var DATA_FATO= SheetBanco.getRange(Linha, 23).getValue();
    var H_FATO= SheetBanco.getRange(Linha, 26).getValue();
    var MUN= SheetBanco.getRange(Linha, 28).getValue();
    var A_MUN= SheetBanco.getRange(Linha, 39).getValue();
    var LOG= SheetBanco.getRange(Linha, 40).getValue();
    var NUM= SheetBanco.getRange(Linha, 41).getValue();
    var COMP= SheetBanco.getRange(Linha, 42).getValue();
    var CEP= SheetBanco.getRange(Linha, 43).getValue();
    var BAIRRO= SheetBanco.getRange(Linha, 44).getValue();
    var LAT= SheetBanco.getRange(Linha, 45).getValue();
    var LNG= SheetBanco.getRange(Linha, 46).getValue();
    var P_REF= SheetBanco.getRange(Linha, 47).getValue();
    var TIPO_LOCAL= SheetBanco.getRange(Linha, 48).getValue();
    var M_UTIL_AUT= SheetBanco.getRange(Linha, 49).getValue();
    var RECUR= SheetBanco.getRange(Linha, 50).getValue();
    var NUM_AGRE= SheetBanco.getRange(Linha, 51).getValue();
    var M_AUTOR= SheetBanco.getRange(Linha, 52).getValue();
    var OBJ= SheetBanco.getRange(Linha, 53).getValue();
    var ADTNT= SheetBanco.getRange(Linha, 54).getValue();
    var MOT= SheetBanco.getRange(Linha, 55).getValue();
    var CADAVER= SheetBanco.getRange(Linha, 56).getValue();
    var R_VIT_AUT= SheetBanco.getRange(Linha, 57).getValue();
    var MID= SheetBanco.getRange(Linha, 58).getValue();
    var INFORMANTE= SheetBanco.getRange(Linha, 59).getValue();
    var N_INFORMANTE= SheetBanco.getRange(Linha, 60).getValue();
    var HIST= SheetBanco.getRange(Linha, 61).getValue();
    var OBS= SheetBanco.getRange(Linha, 62).getValue();
    var N_VIT= SheetBanco.getRange(Linha, 63).getValue();
    var RG_VIT= SheetBanco.getRange(Linha, 64).getValue();
    var CPF_VIT= SheetBanco.getRange(Linha, 65).getValue();
    var SX_VIT= SheetBanco.getRange(Linha, 66).getValue();
    var DN_VIT= SheetBanco.getRange(Linha, 67).getValue();
    var ORIEN_SX_VIT= SheetBanco.getRange(Linha, 71).getValue();
    var COR_VIT= SheetBanco.getRange(Linha, 72).getValue();
    var EC_VIT= SheetBanco.getRange(Linha, 73).getValue();
    var UE_VIT= SheetBanco.getRange(Linha, 74).getValue();
    var FIL_VIT= SheetBanco.getRange(Linha, 75).getValue();
    var NAT_VIT= SheetBanco.getRange(Linha, 76).getValue();
    var NAC_VIT= SheetBanco.getRange(Linha, 77).getValue();
    var COND_FIS_VIT= SheetBanco.getRange(Linha, 78).getValue();
    var ALCUNHA_VIT= SheetBanco.getRange(Linha, 79).getValue();
    var TAT_VIT= SheetBanco.getRange(Linha, 80).getValue();
    var ESC_VIT= SheetBanco.getRange(Linha, 81).getValue();
    //var nome_var = SheetBanco.getRange(Linha, col).getValue();
    var END_RES_VIT= SheetBanco.getRange(Linha, 82).getValue();
    var MUN_VIT= SheetBanco.getRange(Linha, 83).getValue();
    var BAIRRO_VIT= SheetBanco.getRange(Linha, 84).getValue();
    var PROF_VIT= SheetBanco.getRange(Linha, 85).getValue();
    var END_PROF_VIT= SheetBanco.getRange(Linha, 86).getValue();
    var MUN_PROF_VIT= SheetBanco.getRange(Linha, 87).getValue();
    var BAIRRO_PROF_VIT= SheetBanco.getRange(Linha, 88).getValue();
    var N_PAI_VIT= SheetBanco.getRange(Linha, 89).getValue();
    var ESC_PAI_VIT= SheetBanco.getRange(Linha, 90).getValue();
    var HIS_AC_PAI_VIT= SheetBanco.getRange(Linha, 91).getValue();
    var HIS_AC_PAI_VDOM_VIT= SheetBanco.getRange(Linha, 92).getValue();
    //var nome_var = SheetBanco.getRange(Linha, col).getValue();
    var N_MAE_VIT= SheetBanco.getRange(Linha, 93).getValue();
    var ESC_MAEVIT= SheetBanco.getRange(Linha, 94).getValue();
    var HIS_AC_MAE_VIT= SheetBanco.getRange(Linha, 95).getValue();
    var VIOL_DOM_MAE_VIT= SheetBanco.getRange(Linha, 96).getValue();
    //var nome_var = SheetBanco.getRange(Linha, col).getValue();
    var MP_VIT= SheetBanco.getRange(Linha, 97).getValue();
    var MI_VIT= SheetBanco.getRange(Linha, 98).getValue();
    var HIS_MI_VIT= SheetBanco.getRange(Linha, 99).getValue();
    var ENVTRAF_VIT= SheetBanco.getRange(Linha, 100).getValue();
    var FOR_VIT= SheetBanco.getRange(Linha, 101).getValue();
    var ANTEPOL_VIT= SheetBanco.getRange(Linha, 102).getValue();
    var ANTECRIM_VIT= SheetBanco.getRange(Linha, 103).getValue();
    //var nome_var = SheetBanco.getRange(Linha, col).getValue();
    //var nome_var = SheetBanco.getRange(Linha, col).getValue();
    var ANTEC_TRANS_VIT= SheetBanco.getRange(Linha, 104).getValue();
    var ANTEC_DIR_HAB_VIT= SheetBanco.getRange(Linha, 105).getValue();
    var ANTEC_FUG_RECAP_VIT= SheetBanco.getRange(Linha, 106).getValue();
    var ANTEC_ENT_POS_VIT= SheetBanco.getRange(Linha, 107).getValue();
    var ANTEC_ENT_TRAF_VIT= SheetBanco.getRange(Linha, 108).getValue();
    var ANTEC_PORT_ARMAS_VIT= SheetBanco.getRange(Linha, 109).getValue();
    var ANTEC_RECEP_VIT= SheetBanco.getRange(Linha, 110).getValue();
    var ANTEC_FURTO_VIT= SheetBanco.getRange(Linha, 111).getValue();
    var ANTEC_ROUBO_PED_VIT= SheetBanco.getRange(Linha, 112).getValue();
    var ANTEC_ROUBO_BANC_VIT= SheetBanco.getRange(Linha, 113).getValue();
    var ANTEC_ROUBO_TRANSP_PUB_VIT= SheetBanco.getRange(Linha, 114).getValue();
    var ANTEC_ROUBO_TRANSP_IND_VIT= SheetBanco.getRange(Linha, 115).getValue();
    var ANTEC_ROUBO_TRANSP_IND_APP_VIT= SheetBanco.getRange(Linha, 116).getValue();
    var ANTEC_ROUBO_TRANSP_VEIC_VIT= SheetBanco.getRange(Linha, 117).getValue();
    var ANTEC_ROUBO_RESID_VIT= SheetBanco.getRange(Linha, 118).getValue();
    var ANTEC_ROUBO_OUT_VIT= SheetBanco.getRange(Linha, 119).getValue();
    var ANTEC_LATROC_VIT= SheetBanco.getRange(Linha, 120).getValue();
    var ANTEC_AME_VIT= SheetBanco.getRange(Linha, 121).getValue();
    var ANTEC_LES_CORP_VIT= SheetBanco.getRange(Linha, 122).getValue();
    var ANTEC_LES_CORP_MOR_VIT= SheetBanco.getRange(Linha, 123).getValue();
    var ANTEC_HOM_VIT= SheetBanco.getRange(Linha, 124).getValue();
    var ANTEC_MP_VIT= SheetBanco.getRange(Linha, 125).getValue();
    var ANTEC_CRIM_SEX_VIT= SheetBanco.getRange(Linha, 126).getValue();
    var ANTEC_ESTUP_VUL_VIT= SheetBanco.getRange(Linha, 127).getValue();
    var ANTEC_OUT_VIT= SheetBanco.getRange(Linha, 128).getValue();
    var HIST_VIT_TRANS_VIT= SheetBanco.getRange(Linha, 129).getValue();
    var HIST_VIT_FURTO_VIT= SheetBanco.getRange(Linha, 130).getValue();
    var HIST_VIT_ROUBO_PED_VIT= SheetBanco.getRange(Linha, 131).getValue();
    var HIST_VIT_ROUBO_BANC_VIT= SheetBanco.getRange(Linha, 132).getValue();
    var HIST_VIT_ROUBO_TRANSP_PUB_VIT= SheetBanco.getRange(Linha, 133).getValue();
    var HIST_VIT_ROUBO_TRANSP_IND_VIT= SheetBanco.getRange(Linha, 134).getValue();
    var HIST_VIT_ROUBO_TRANSP_IND_APP_VIT= SheetBanco.getRange(Linha, 135).getValue();
    var HIST_VIT_ROUBO_TRANSP_VEIC_VIT= SheetBanco.getRange(Linha, 136).getValue();
    var HIST_VIT_ROUBO_RESID_VIT= SheetBanco.getRange(Linha, 137).getValue();
    var HIST_VIT_ROUBO_OUT_VIT= SheetBanco.getRange(Linha, 138).getValue();
    var HIST_VIT_LATROC_VIT= SheetBanco.getRange(Linha, 139).getValue();
    var HIST_VIT_AME_VIT= SheetBanco.getRange(Linha, 140).getValue();
    var HIST_VIT_LES_CORP_VIT= SheetBanco.getRange(Linha, 141).getValue();
    var HIST_VIT_HOM_VIT= SheetBanco.getRange(Linha, 142).getValue();
    var HIST_VIT_MP_VIT= SheetBanco.getRange(Linha, 143).getValue();
    var HIST_VIT_CRIM_SEX_VIT= SheetBanco.getRange(Linha, 144).getValue();
    var HIST_VIT_ESTUP_VUL_VIT= SheetBanco.getRange(Linha, 145).getValue();
    var HIST_VIT_OUT_VIT= SheetBanco.getRange(Linha, 146).getValue();
    //var nome_var = SheetBanco.getRange(Linha, col).getValue();
    //var nome_var = SheetBanco.getRange(Linha, col).getValue();
    //var nome_var = SheetBanco.getRange(Linha, col).getValue();
    var N_ACU_1= SheetBanco.getRange(Linha, 147).getValue();
    var RG_ACU_1= SheetBanco.getRange(Linha, 148).getValue();
    var CPF_ACU_1= SheetBanco.getRange(Linha, 149).getValue();
    var SX_ACU_1= SheetBanco.getRange(Linha, 150).getValue();
    var DN_ACU_1= SheetBanco.getRange(Linha, 151).getValue();
    var ORIEN_SX_ACU_1= SheetBanco.getRange(Linha, 155).getValue();
    var COR_ACU_1= SheetBanco.getRange(Linha, 156).getValue();
    var EC_ACU_1= SheetBanco.getRange(Linha, 157).getValue();
    var UE_ACU_1= SheetBanco.getRange(Linha, 158).getValue();
    var FIL_ACU_1= SheetBanco.getRange(Linha, 159).getValue();
    var NAT_ACU_1= SheetBanco.getRange(Linha, 160).getValue();
    var NAC_ACU_1= SheetBanco.getRange(Linha, 161).getValue();
    var COND_FIS_ACU_1= SheetBanco.getRange(Linha, 162).getValue();
    var ALCUNHA_ACU_1= SheetBanco.getRange(Linha, 163).getValue();
    var TAT_ACU_1= SheetBanco.getRange(Linha, 164).getValue();
    var ESC_ACU_1= SheetBanco.getRange(Linha, 165).getValue();
    //var nome_var = SheetBanco.getRange(Linha, col).getValue();
    var END_RES_ACU_1= SheetBanco.getRange(Linha, 166).getValue();
    var MUN_ACU_1= SheetBanco.getRange(Linha, 167).getValue();
    var BAIRRO_ACU_1= SheetBanco.getRange(Linha, 168).getValue();
    var PROF_ACU_1= SheetBanco.getRange(Linha, 169).getValue();
    var END_PROF_ACU_1= SheetBanco.getRange(Linha, 170).getValue();
    var MUN_PROF_ACU_1= SheetBanco.getRange(Linha, 171).getValue();
    var BAIRRO_PROF_ACU_1= SheetBanco.getRange(Linha, 172).getValue();
    var N_PAI_ACU_1= SheetBanco.getRange(Linha, 173).getValue();
    var ESC_PAI_ACU_1= SheetBanco.getRange(Linha, 174).getValue();
    var HIS_AC_PAI_ACU_1= SheetBanco.getRange(Linha, 175).getValue();
    var HIS_AC_PAI_VDOM_ACU_1= SheetBanco.getRange(Linha, 176).getValue();
    //var nome_var = SheetBanco.getRange(Linha, col).getValue();
    var ESC_MAE_ACU_1= SheetBanco.getRange(Linha, 177).getValue();
    var N_MAE_ACU_1= SheetBanco.getRange(Linha, 178).getValue();
    var HIS_AC_MAE_ACU_1= SheetBanco.getRange(Linha, 179).getValue();
    var VIOL_DOM_MAE_ACU_1= SheetBanco.getRange(Linha, 180).getValue();
    //var nome_var = SheetBanco.getRange(Linha, col).getValue();
    var MP_ACU_1= SheetBanco.getRange(Linha, 181).getValue();
    var MI_ACU_1= SheetBanco.getRange(Linha, 182).getValue();
    var HIS_MI_ACU_1= SheetBanco.getRange(Linha, 183).getValue();
    var ENVTRAF_ACU_1= SheetBanco.getRange(Linha, 184).getValue();
    var FOR_ACU_1= SheetBanco.getRange(Linha, 185).getValue();
    var ANTEPOL_ACU_1= SheetBanco.getRange(Linha, 186).getValue();
    var ANTECRIM_ACU_1= SheetBanco.getRange(Linha, 187).getValue();
    //var nome_var = SheetBanco.getRange(Linha, col).getValue();
    //var nome_var = SheetBanco.getRange(Linha, col).getValue();
    var ANTEC_TRANS_ACU_1= SheetBanco.getRange(Linha, 188).getValue();
    var ANTEC_DIR_HAB_ACU_1= SheetBanco.getRange(Linha, 189).getValue();
    var ANTEC_FUG_RECAP_ACU_1= SheetBanco.getRange(Linha, 190).getValue();
    var ANTEC_ENT_POS_ACU_1= SheetBanco.getRange(Linha, 191).getValue();
    var ANTEC_ENT_TRAF_ACU_1= SheetBanco.getRange(Linha, 192).getValue();
    var ANTEC_PORT_ARMAS_ACU_1= SheetBanco.getRange(Linha, 193).getValue();
    var ANTEC_RECEP_ACU_1= SheetBanco.getRange(Linha, 194).getValue();
    var ANTEC_FURTO_ACU_1= SheetBanco.getRange(Linha, 195).getValue();
    var ANTEC_ROUBO_PED_ACU_1= SheetBanco.getRange(Linha, 196).getValue();
    var ANTEC_ROUBO_BANC_ACU_1= SheetBanco.getRange(Linha, 197).getValue();
    var ANTEC_ROUBO_TRANSP_PUB_ACU_1= SheetBanco.getRange(Linha, 198).getValue();
    var ANTEC_ROUBO_TRANSP_IND_ACU_1= SheetBanco.getRange(Linha, 199).getValue();
    var ANTEC_ROUBO_TRANSP_IND_APP_ACU_1= SheetBanco.getRange(Linha, 200).getValue();
    var ANTEC_ROUBO_TRANSP_VEIC_ACU_1= SheetBanco.getRange(Linha, 201).getValue();
    var ANTEC_ROUBO_RESID_ACU_1= SheetBanco.getRange(Linha, 202).getValue();
    var ANTEC_ROUBO_OUT_ACU_1= SheetBanco.getRange(Linha, 203).getValue();
    var ANTEC_LATROC_ACU_1= SheetBanco.getRange(Linha, 204).getValue();
    var ANTEC_AME_ACU_1= SheetBanco.getRange(Linha, 205).getValue();
    var ANTEC_LES_CORP_ACU_1= SheetBanco.getRange(Linha, 206).getValue();
    var ANTEC_LES_CORP_MOR_ACU_1= SheetBanco.getRange(Linha, 207).getValue();
    var ANTEC_HOM_ACU_1= SheetBanco.getRange(Linha, 208).getValue();
    var ANTEC_MP_ACU_1= SheetBanco.getRange(Linha, 209).getValue();
    var ANTEC_CRIM_SEX_ACU_1= SheetBanco.getRange(Linha, 210).getValue();
    var ANTEC_ESTUP_VUL_ACU_1= SheetBanco.getRange(Linha, 211).getValue();
    var ANTEC_OUT_ACU_1= SheetBanco.getRange(Linha, 212).getValue();
    var HIST_VIT_TRANS_ACU_1= SheetBanco.getRange(Linha, 213).getValue();
    var HIST_VIT_FURTO_ACU_1= SheetBanco.getRange(Linha, 214).getValue();
    var HIST_VIT_ROUBO_PED_ACU_1= SheetBanco.getRange(Linha, 215).getValue();
    var HIST_VIT_ROUBO_BANC_ACU_1= SheetBanco.getRange(Linha, 216).getValue();
    var HIST_VIT_ROUBO_TRANSP_PUB_ACU_1= SheetBanco.getRange(Linha, 217).getValue();
    var HIST_VIT_ROUBO_TRANSP_IND_ACU_1= SheetBanco.getRange(Linha, 218).getValue();
    var HIST_VIT_ROUBO_TRANSP_IND_APP_ACU_1= SheetBanco.getRange(Linha, 219).getValue();
    var HIST_VIT_ROUBO_TRANSP_VEIC_ACU_1= SheetBanco.getRange(Linha, 220).getValue();
    var HIST_VIT_ROUBO_RESID_ACU_1= SheetBanco.getRange(Linha, 221).getValue();
    var HIST_VIT_ROUBO_OUT_ACU_1= SheetBanco.getRange(Linha, 222).getValue();
    var HIST_VIT_LATROC_ACU_1= SheetBanco.getRange(Linha, 223).getValue();
    var HIST_VIT_AME_ACU_1= SheetBanco.getRange(Linha, 224).getValue();
    var HIST_VIT_LES_CORP_ACU_1= SheetBanco.getRange(Linha, 225).getValue();
    var HIST_VIT_HOM_ACU_1= SheetBanco.getRange(Linha, 226).getValue();
    var HIST_VIT_MP_ACU_1= SheetBanco.getRange(Linha, 227).getValue();
    var HIST_VIT_CRIM_SEX_ACU_1= SheetBanco.getRange(Linha, 228).getValue();
    var HIST_VIT_ESTUP_VUL_ACU_1= SheetBanco.getRange(Linha, 229).getValue();
    var HIST_VIT_OUT_ACU_1= SheetBanco.getRange(Linha, 230).getValue();
    //var nome_var = SheetBanco.getRange(Linha, col).getValue();
    //var nome_var = SheetBanco.getRange(Linha, col).getValue();
    //var nome_var = SheetBanco.getRange(Linha, col).getValue();
    var N_ACU_2= SheetBanco.getRange(Linha, 231).getValue();
    var RG_ACU_2= SheetBanco.getRange(Linha, 232).getValue();
    var CPF_ACU_2= SheetBanco.getRange(Linha, 233).getValue();
    var SX_ACU_2= SheetBanco.getRange(Linha, 234).getValue();
    var DN_ACU_2= SheetBanco.getRange(Linha, 235).getValue();
    var ORIEN_SX_ACU_2= SheetBanco.getRange(Linha, 239).getValue();
    var COR_ACU_2= SheetBanco.getRange(Linha, 240).getValue();
    var EC_ACU_2= SheetBanco.getRange(Linha, 241).getValue();
    var UE_ACU_2= SheetBanco.getRange(Linha, 242).getValue();
    var FIL_ACU_2= SheetBanco.getRange(Linha, 243).getValue();
    var NAT_ACU_2= SheetBanco.getRange(Linha, 244).getValue();
    var NAC_ACU_2= SheetBanco.getRange(Linha, 245).getValue();
    var COND_FIS_ACU_2= SheetBanco.getRange(Linha, 246).getValue();
    var ALCUNHA_ACU_2= SheetBanco.getRange(Linha, 247).getValue();
    var TAT_ACU_2= SheetBanco.getRange(Linha, 248).getValue();
    var ESC_ACU_2= SheetBanco.getRange(Linha, 249).getValue();
    //var nome_var = SheetBanco.getRange(Linha, col).getValue();
    var END_RES_ACU_2= SheetBanco.getRange(Linha, 250).getValue();
    var MUN_ACU_2= SheetBanco.getRange(Linha, 251).getValue();
    var BAIRRO_ACU_2= SheetBanco.getRange(Linha, 252).getValue();
    var PROF_ACU_2= SheetBanco.getRange(Linha, 253).getValue();
    var END_PROF_ACU_2= SheetBanco.getRange(Linha, 254).getValue();
    var MUN_PROF_ACU_2= SheetBanco.getRange(Linha, 255).getValue();
    var BAIRRO_PROF_ACU_2= SheetBanco.getRange(Linha, 256).getValue();
    var N_PAI_ACU_2= SheetBanco.getRange(Linha, 257).getValue();
    var ESC_PAI_ACU_2= SheetBanco.getRange(Linha, 258).getValue();
    var HIS_AC_PAI_ACU_2= SheetBanco.getRange(Linha, 259).getValue();
    var HIS_AC_PAI_VDOM_ACU_2= SheetBanco.getRange(Linha, 260).getValue();
    //var nome_var = SheetBanco.getRange(Linha, col).getValue();
    var ESC_MAE_ACU_2= SheetBanco.getRange(Linha, 261).getValue();
    var N_MAE_ACU_2= SheetBanco.getRange(Linha, 262).getValue();
    var HIS_AC_MAE_ACU_2= SheetBanco.getRange(Linha, 263).getValue();
    var VIOL_DOM_MAE_ACU_2= SheetBanco.getRange(Linha, 264).getValue();
    //var nome_var = SheetBanco.getRange(Linha, col).getValue();
    var MP_ACU_2= SheetBanco.getRange(Linha, 265).getValue();
    var MI_ACU_2= SheetBanco.getRange(Linha, 266).getValue();
    var HIS_MI_ACU_2= SheetBanco.getRange(Linha, 267).getValue();
    var ENVTRAF_ACU_2= SheetBanco.getRange(Linha, 268).getValue();
    var FOR_ACU_2= SheetBanco.getRange(Linha, 269).getValue();
    var ANTEPOL_ACU_2= SheetBanco.getRange(Linha, 270).getValue();
    var ANTECRIM_ACU_2= SheetBanco.getRange(Linha, 271).getValue();
    //var nome_var = SheetBanco.getRange(Linha, col).getValue();
    //var nome_var = SheetBanco.getRange(Linha, col).getValue();
    var ANTEC_TRANS_ACU_2= SheetBanco.getRange(Linha, 272).getValue();
    var ANTEC_DIR_HAB_ACU_2= SheetBanco.getRange(Linha, 273).getValue();
    var ANTEC_FUG_RECAP_ACU_2= SheetBanco.getRange(Linha, 274).getValue();
    var ANTEC_ENT_POS_ACU_2= SheetBanco.getRange(Linha, 275).getValue();
    var ANTEC_ENT_TRAF_ACU_2= SheetBanco.getRange(Linha, 276).getValue();
    var ANTEC_PORT_ARMAS_ACU_2= SheetBanco.getRange(Linha, 277).getValue();
    var ANTEC_RECEP_ACU_2= SheetBanco.getRange(Linha, 278).getValue();
    var ANTEC_FURTO_ACU_2= SheetBanco.getRange(Linha, 279).getValue();
    var ANTEC_ROUBO_PED_ACU_2= SheetBanco.getRange(Linha, 280).getValue();
    var ANTEC_ROUBO_BANC_ACU_2= SheetBanco.getRange(Linha, 281).getValue();
    var ANTEC_ROUBO_TRANSP_PUB_ACU_2= SheetBanco.getRange(Linha, 282).getValue();
    var ANTEC_ROUBO_TRANSP_IND_ACU_2= SheetBanco.getRange(Linha, 283).getValue();
    var ANTEC_ROUBO_TRANSP_IND_APP_ACU_2= SheetBanco.getRange(Linha, 284).getValue();
    var ANTEC_ROUBO_TRANSP_VEIC_ACU_2= SheetBanco.getRange(Linha, 285).getValue();
    var ANTEC_ROUBO_RESID_ACU_2= SheetBanco.getRange(Linha, 286).getValue();
    var ANTEC_ROUBO_OUT_ACU_2= SheetBanco.getRange(Linha, 287).getValue();
    var ANTEC_LATROC_ACU_2= SheetBanco.getRange(Linha, 288).getValue();
    var ANTEC_AME_ACU_2= SheetBanco.getRange(Linha, 289).getValue();
    var ANTEC_LES_CORP_ACU_2= SheetBanco.getRange(Linha, 290).getValue();
    var ANTEC_LES_CORP_MOR_ACU_2= SheetBanco.getRange(Linha, 291).getValue();
    var ANTEC_HOM_ACU_2= SheetBanco.getRange(Linha, 292).getValue();
    var ANTEC_MP_ACU_2= SheetBanco.getRange(Linha, 293).getValue();
    var ANTEC_CRIM_SEX_ACU_2= SheetBanco.getRange(Linha, 294).getValue();
    var ANTEC_ESTUP_VUL_ACU_2= SheetBanco.getRange(Linha, 295).getValue();
    var ANTEC_OUT_ACU_2= SheetBanco.getRange(Linha, 296).getValue();
    var HIST_VIT_TRANS_ACU_2= SheetBanco.getRange(Linha, 297).getValue();
    var HIST_VIT_FURTO_ACU_2= SheetBanco.getRange(Linha, 298).getValue();
    var HIST_VIT_ROUBO_PED_ACU_2= SheetBanco.getRange(Linha, 299).getValue();
    var HIST_VIT_ROUBO_BANC_ACU_2= SheetBanco.getRange(Linha, 300).getValue();
    var HIST_VIT_ROUBO_TRANSP_PUB_ACU_2= SheetBanco.getRange(Linha, 301).getValue();
    var HIST_VIT_ROUBO_TRANSP_IND_ACU_2= SheetBanco.getRange(Linha, 302).getValue();
    var HIST_VIT_ROUBO_TRANSP_IND_APP_ACU_2= SheetBanco.getRange(Linha, 303).getValue();
    var HIST_VIT_ROUBO_TRANSP_VEIC_ACU_2= SheetBanco.getRange(Linha, 304).getValue();
    var HIST_VIT_ROUBO_RESID_ACU_2= SheetBanco.getRange(Linha, 305).getValue();
    var HIST_VIT_ROUBO_OUT_ACU_2= SheetBanco.getRange(Linha, 306).getValue();
    var HIST_VIT_LATROC_ACU_2= SheetBanco.getRange(Linha, 307).getValue();
    var HIST_VIT_AME_ACU_2= SheetBanco.getRange(Linha, 308).getValue();
    var HIST_VIT_LES_CORP_ACU_2= SheetBanco.getRange(Linha, 309).getValue();
    var HIST_VIT_HOM_ACU_2= SheetBanco.getRange(Linha, 310).getValue();
    var HIST_VIT_MP_ACU_2= SheetBanco.getRange(Linha, 311).getValue();
    var HIST_VIT_CRIM_SEX_ACU_2= SheetBanco.getRange(Linha, 312).getValue();
    var HIST_VIT_ESTUP_VUL_ACU_2= SheetBanco.getRange(Linha, 313).getValue();
    var HIST_VIT_OUT_ACU_2= SheetBanco.getRange(Linha, 314).getValue();
    //var nome_var = SheetBanco.getRange(Linha, col).getValue();
    //var nome_var = SheetBanco.getRange(Linha, col).getValue();
    //var nome_var = SheetBanco.getRange(Linha, col).getValue();
    var N_ACU_3= SheetBanco.getRange(Linha, 315).getValue();
    var RG_ACU_3= SheetBanco.getRange(Linha, 316).getValue();
    var CPF_ACU_3= SheetBanco.getRange(Linha, 317).getValue();
    var SX_ACU_3= SheetBanco.getRange(Linha, 318).getValue();
    var DN_ACU_3= SheetBanco.getRange(Linha, 319).getValue();
    var ORIEN_SX_ACU_3= SheetBanco.getRange(Linha, 323).getValue();
    var COR_ACU_3= SheetBanco.getRange(Linha, 324).getValue();
    var EC_ACU_3= SheetBanco.getRange(Linha, 325).getValue();
    var UE_ACU_3= SheetBanco.getRange(Linha, 326).getValue();
    var FIL_ACU_3= SheetBanco.getRange(Linha, 327).getValue();
    var NAT_ACU_3= SheetBanco.getRange(Linha, 328).getValue();
    var NAC_ACU_3= SheetBanco.getRange(Linha, 329).getValue();
    var COND_FIS_ACU_3= SheetBanco.getRange(Linha, 330).getValue();
    var ALCUNHA_ACU_3= SheetBanco.getRange(Linha, 331).getValue();
    var TAT_ACU_3= SheetBanco.getRange(Linha, 332).getValue();
    var ESC_ACU_3= SheetBanco.getRange(Linha, 333).getValue();
    //var nome_var = SheetBanco.getRange(Linha, col).getValue();
    var END_RES_ACU_3= SheetBanco.getRange(Linha, 334).getValue();
    var MUN_ACU_3= SheetBanco.getRange(Linha, 335).getValue();
    var BAIRRO_ACU_3= SheetBanco.getRange(Linha, 336).getValue();
    var PROF_ACU_3= SheetBanco.getRange(Linha, 337).getValue();
    var END_PROF_ACU_3= SheetBanco.getRange(Linha, 338).getValue();
    var MUN_PROF_ACU_3= SheetBanco.getRange(Linha, 339).getValue();
    var BAIRRO_PROF_ACU_3= SheetBanco.getRange(Linha, 340).getValue();
    var N_PAI_ACU_3= SheetBanco.getRange(Linha, 341).getValue();
    var ESC_PAI_ACU_3= SheetBanco.getRange(Linha, 342).getValue();
    var HIS_AC_PAI_ACU_3= SheetBanco.getRange(Linha, 343).getValue();
    var HIS_AC_PAI_VDOM_ACU_3= SheetBanco.getRange(Linha, 344).getValue();
    //var nome_var = SheetBanco.getRange(Linha, col).getValue();
    var ESC_MAE_ACU_3= SheetBanco.getRange(Linha, 345).getValue();
    var N_MAE_ACU_3= SheetBanco.getRange(Linha, 346).getValue();
    var HIS_AC_MAE_ACU_3= SheetBanco.getRange(Linha, 347).getValue();
    var VIOL_DOM_MAE_ACU_3= SheetBanco.getRange(Linha, 348).getValue();
    //var nome_var = SheetBanco.getRange(Linha, col).getValue();
    var MP_ACU_3= SheetBanco.getRange(Linha, 349).getValue();
    var MI_ACU_3= SheetBanco.getRange(Linha, 350).getValue();
    var HIS_MI_ACU_3= SheetBanco.getRange(Linha, 351).getValue();
    var ENVTRAF_ACU_3= SheetBanco.getRange(Linha, 352).getValue();
    var FOR_ACU_3= SheetBanco.getRange(Linha, 353).getValue();
    var ANTEPOL_ACU_3= SheetBanco.getRange(Linha, 354).getValue();
    var ANTECRIM_ACU_3= SheetBanco.getRange(Linha, 355).getValue();
    //var nome_var = SheetBanco.getRange(Linha, col).getValue();
    //var nome_var = SheetBanco.getRange(Linha, col).getValue();
    var ANTEC_TRANS_ACU_3= SheetBanco.getRange(Linha, 356).getValue();
    var ANTEC_DIR_HAB_ACU_3= SheetBanco.getRange(Linha, 357).getValue();
    var ANTEC_FUG_RECAP_ACU_3= SheetBanco.getRange(Linha, 358).getValue();
    var ANTEC_ENT_POS_ACU_3= SheetBanco.getRange(Linha, 359).getValue();
    var ANTEC_ENT_TRAF_ACU_3= SheetBanco.getRange(Linha, 360).getValue();
    var ANTEC_PORT_ARMAS_ACU_3= SheetBanco.getRange(Linha, 361).getValue();
    var ANTEC_RECEP_ACU_3= SheetBanco.getRange(Linha, 362).getValue();
    var ANTEC_FURTO_ACU_3= SheetBanco.getRange(Linha, 363).getValue();
    var ANTEC_ROUBO_PED_ACU_3= SheetBanco.getRange(Linha, 364).getValue();
    var ANTEC_ROUBO_BANC_ACU_3= SheetBanco.getRange(Linha, 365).getValue();
    var ANTEC_ROUBO_TRANSP_PUB_ACU_3= SheetBanco.getRange(Linha, 366).getValue();
    var ANTEC_ROUBO_TRANSP_IND_ACU_3= SheetBanco.getRange(Linha, 367).getValue();
    var ANTEC_ROUBO_TRANSP_IND_APP_ACU_3= SheetBanco.getRange(Linha, 368).getValue();
    var ANTEC_ROUBO_TRANSP_VEIC_ACU_3= SheetBanco.getRange(Linha, 369).getValue();
    var ANTEC_ROUBO_RESID_ACU_3= SheetBanco.getRange(Linha, 370).getValue();
    var ANTEC_ROUBO_OUT_ACU_3= SheetBanco.getRange(Linha, 371).getValue();
    var ANTEC_LATROC_ACU_3= SheetBanco.getRange(Linha, 372).getValue();
    var ANTEC_AME_ACU_3= SheetBanco.getRange(Linha, 373).getValue();
    var ANTEC_LES_CORP_ACU_3= SheetBanco.getRange(Linha, 374).getValue();
    var ANTEC_LES_CORP_MOR_ACU_3= SheetBanco.getRange(Linha, 375).getValue();
    var ANTEC_HOM_ACU_3= SheetBanco.getRange(Linha, 376).getValue();
    var ANTEC_MP_ACU_3= SheetBanco.getRange(Linha, 377).getValue();
    var ANTEC_CRIM_SEX_ACU_3= SheetBanco.getRange(Linha, 378).getValue();
    var ANTEC_ESTUP_VUL_ACU_3= SheetBanco.getRange(Linha, 379).getValue();
    var ANTEC_OUT_ACU_3= SheetBanco.getRange(Linha, 380).getValue();
    var HIST_VIT_TRANS_ACU_3= SheetBanco.getRange(Linha, 381).getValue();
    var HIST_VIT_FURTO_ACU_3= SheetBanco.getRange(Linha, 382).getValue();
    var HIST_VIT_ROUBO_PED_ACU_3= SheetBanco.getRange(Linha, 383).getValue();
    var HIST_VIT_ROUBO_BANC_ACU_3= SheetBanco.getRange(Linha, 384).getValue();
    var HIST_VIT_ROUBO_TRANSP_PUB_ACU_3= SheetBanco.getRange(Linha, 385).getValue();
    var HIST_VIT_ROUBO_TRANSP_IND_ACU_3= SheetBanco.getRange(Linha, 386).getValue();
    var HIST_VIT_ROUBO_TRANSP_IND_APP_ACU_3= SheetBanco.getRange(Linha, 387).getValue();
    var HIST_VIT_ROUBO_TRANSP_VEIC_ACU_3= SheetBanco.getRange(Linha, 388).getValue();
    var HIST_VIT_ROUBO_RESID_ACU_3= SheetBanco.getRange(Linha, 389).getValue();
    var HIST_VIT_ROUBO_OUT_ACU_3= SheetBanco.getRange(Linha, 390).getValue();
    var HIST_VIT_LATROC_ACU_3= SheetBanco.getRange(Linha, 391).getValue();
    var HIST_VIT_AME_ACU_3= SheetBanco.getRange(Linha, 392).getValue();
    var HIST_VIT_LES_CORP_ACU_3= SheetBanco.getRange(Linha, 393).getValue();
    var HIST_VIT_HOM_ACU_3= SheetBanco.getRange(Linha, 394).getValue();
    var HIST_VIT_MP_ACU_3= SheetBanco.getRange(Linha, 395).getValue();
    var HIST_VIT_CRIM_SEX_ACU_3= SheetBanco.getRange(Linha, 396).getValue();
    var HIST_VIT_ESTUP_VUL_ACU_3= SheetBanco.getRange(Linha, 397).getValue();
    var HIST_VIT_OUT_ACU_3= SheetBanco.getRange(Linha, 398).getValue();
    //var nome_var = SheetBanco.getRange(Linha, col).getValue();
    //var nome_var = SheetBanco.getRange(Linha, col).getValue();
    //var nome_var = SheetBanco.getRange(Linha, col).getValue();
    var N_ACU_4= SheetBanco.getRange(Linha, 399).getValue();
    var RG_ACU_4= SheetBanco.getRange(Linha, 400).getValue();
    var CPF_ACU_4= SheetBanco.getRange(Linha, 401).getValue();
    var SX_ACU_4= SheetBanco.getRange(Linha, 402).getValue();
    var DN_ACU_4= SheetBanco.getRange(Linha, 403).getValue();
    var ORIEN_SX_ACU_4= SheetBanco.getRange(Linha, 407).getValue();
    var COR_ACU_4= SheetBanco.getRange(Linha, 408).getValue();
    var EC_ACU_4= SheetBanco.getRange(Linha, 409).getValue();
    var UE_ACU_4= SheetBanco.getRange(Linha, 410).getValue();
    var FIL_ACU_4= SheetBanco.getRange(Linha, 411).getValue();
    var NAT_ACU_4= SheetBanco.getRange(Linha, 412).getValue();
    var NAC_ACU_4= SheetBanco.getRange(Linha, 413).getValue();
    var COND_FIS_ACU_4= SheetBanco.getRange(Linha, 414).getValue();
    var ALCUNHA_ACU_4= SheetBanco.getRange(Linha, 415).getValue();
    var TAT_ACU_4= SheetBanco.getRange(Linha, 416).getValue();
    var ESC_ACU_4= SheetBanco.getRange(Linha, 417).getValue();
    //var nome_var = SheetBanco.getRange(Linha, col).getValue();
    var END_RES_ACU_4= SheetBanco.getRange(Linha, 418).getValue();
    var MUN_ACU_4= SheetBanco.getRange(Linha, 419).getValue();
    var BAIRRO_ACU_4= SheetBanco.getRange(Linha, 420).getValue();
    var PROF_ACU_4= SheetBanco.getRange(Linha, 421).getValue();
    var END_PROF_ACU_4= SheetBanco.getRange(Linha, 422).getValue();
    var MUN_PROF_ACU_4= SheetBanco.getRange(Linha, 423).getValue();
    var BAIRRO_PROF_ACU_4= SheetBanco.getRange(Linha, 424).getValue();
    var N_PAI_ACU_4= SheetBanco.getRange(Linha, 425).getValue();
    var ESC_PAI_ACU_4= SheetBanco.getRange(Linha, 426).getValue();
    var HIS_AC_PAI_ACU_4= SheetBanco.getRange(Linha, 427).getValue();
    var HIS_AC_PAI_VDOM_ACU_4= SheetBanco.getRange(Linha, 428).getValue();
    //var nome_var = SheetBanco.getRange(Linha, col).getValue();
    var ESC_MAE_ACU_4= SheetBanco.getRange(Linha, 429).getValue();
    var N_MAE_ACU_4= SheetBanco.getRange(Linha, 430).getValue();
    var HIS_AC_MAE_ACU_4= SheetBanco.getRange(Linha, 431).getValue();
    var VIOL_DOM_MAE_ACU_4= SheetBanco.getRange(Linha, 432).getValue();
    //var nome_var = SheetBanco.getRange(Linha, col).getValue();
    var MP_ACU_4= SheetBanco.getRange(Linha, 433).getValue();
    var MI_ACU_4= SheetBanco.getRange(Linha, 434).getValue();
    var HIS_MI_ACU_4= SheetBanco.getRange(Linha, 435).getValue();
    var ENVTRAF_ACU_4= SheetBanco.getRange(Linha, 436).getValue();
    var FOR_ACU_4= SheetBanco.getRange(Linha, 437).getValue();
    var ANTEPOL_ACU_4= SheetBanco.getRange(Linha, 438).getValue();
    var ANTECRIM_ACU_4= SheetBanco.getRange(Linha, 439).getValue();
    //var nome_var = SheetBanco.getRange(Linha, col).getValue();
    //var nome_var = SheetBanco.getRange(Linha, col).getValue();
    var ANTEC_TRANS_ACU_4= SheetBanco.getRange(Linha, 440).getValue();
    var ANTEC_DIR_HAB_ACU_4= SheetBanco.getRange(Linha, 441).getValue();
    var ANTEC_FUG_RECAP_ACU_4= SheetBanco.getRange(Linha, 442).getValue();
    var ANTEC_ENT_POS_ACU_4= SheetBanco.getRange(Linha, 443).getValue();
    var ANTEC_ENT_TRAF_ACU_4= SheetBanco.getRange(Linha, 444).getValue();
    var ANTEC_PORT_ARMAS_ACU_4= SheetBanco.getRange(Linha, 445).getValue();
    var ANTEC_RECEP_ACU_4= SheetBanco.getRange(Linha, 446).getValue();
    var ANTEC_FURTO_ACU_4= SheetBanco.getRange(Linha, 447).getValue();
    var ANTEC_ROUBO_PED_ACU_4= SheetBanco.getRange(Linha, 448).getValue();
    var ANTEC_ROUBO_BANC_ACU_4= SheetBanco.getRange(Linha, 449).getValue();
    var ANTEC_ROUBO_TRANSP_PUB_ACU_4= SheetBanco.getRange(Linha, 450).getValue();
    var ANTEC_ROUBO_TRANSP_IND_ACU_4= SheetBanco.getRange(Linha, 451).getValue();
    var ANTEC_ROUBO_TRANSP_IND_APP_ACU_4= SheetBanco.getRange(Linha, 452).getValue();
    var ANTEC_ROUBO_TRANSP_VEIC_ACU_4= SheetBanco.getRange(Linha, 453).getValue();
    var ANTEC_ROUBO_RESID_ACU_4= SheetBanco.getRange(Linha, 454).getValue();
    var ANTEC_ROUBO_OUT_ACU_4= SheetBanco.getRange(Linha, 455).getValue();
    var ANTEC_LATROC_ACU_4= SheetBanco.getRange(Linha, 456).getValue();
    var ANTEC_AME_ACU_4= SheetBanco.getRange(Linha, 457).getValue();
    var ANTEC_LES_CORP_ACU_4= SheetBanco.getRange(Linha, 458).getValue();
    var ANTEC_LES_CORP_MOR_ACU_4= SheetBanco.getRange(Linha, 459).getValue();
    var ANTEC_HOM_ACU_4= SheetBanco.getRange(Linha, 460).getValue();
    var ANTEC_MP_ACU_4= SheetBanco.getRange(Linha, 461).getValue();
    var ANTEC_CRIM_SEX_ACU_4= SheetBanco.getRange(Linha, 462).getValue();
    var ANTEC_ESTUP_VUL_ACU_4= SheetBanco.getRange(Linha, 463).getValue();
    var ANTEC_OUT_ACU_4= SheetBanco.getRange(Linha, 464).getValue();
    var HIST_VIT_TRANS_ACU_4= SheetBanco.getRange(Linha, 465).getValue();
    var HIST_VIT_FURTO_ACU_4= SheetBanco.getRange(Linha, 466).getValue();
    var HIST_VIT_ROUBO_PED_ACU_4= SheetBanco.getRange(Linha, 467).getValue();
    var HIST_VIT_ROUBO_BANC_ACU_4= SheetBanco.getRange(Linha, 468).getValue();
    var HIST_VIT_ROUBO_TRANSP_PUB_ACU_4= SheetBanco.getRange(Linha, 469).getValue();
    var HIST_VIT_ROUBO_TRANSP_IND_ACU_4= SheetBanco.getRange(Linha, 470).getValue();
    var HIST_VIT_ROUBO_TRANSP_IND_APP_ACU_4= SheetBanco.getRange(Linha, 471).getValue();
    var HIST_VIT_ROUBO_TRANSP_VEIC_ACU_4= SheetBanco.getRange(Linha, 472).getValue();
    var HIST_VIT_ROUBO_RESID_ACU_4= SheetBanco.getRange(Linha, 473).getValue();
    var HIST_VIT_ROUBO_OUT_ACU_4= SheetBanco.getRange(Linha, 474).getValue();
    var HIST_VIT_LATROC_ACU_4= SheetBanco.getRange(Linha, 475).getValue();
    var HIST_VIT_AME_ACU_4= SheetBanco.getRange(Linha, 476).getValue();
    var HIST_VIT_LES_CORP_ACU_4= SheetBanco.getRange(Linha, 477).getValue();
    var HIST_VIT_HOM_ACU_4= SheetBanco.getRange(Linha, 478).getValue();
    var HIST_VIT_MP_ACU_4= SheetBanco.getRange(Linha, 479).getValue();
    var HIST_VIT_CRIM_SEX_ACU_4= SheetBanco.getRange(Linha, 480).getValue();
    var HIST_VIT_ESTUP_VUL_ACU_4= SheetBanco.getRange(Linha, 481).getValue();
    var HIST_VIT_OUT_ACU_4= SheetBanco.getRange(Linha, 482).getValue();
    //var nome_var = SheetBanco.getRange(Linha, col).getValue();
    //var nome_var = SheetBanco.getRange(Linha, col).getValue();
    //var nome_var = SheetBanco.getRange(Linha, col).getValue();
    var N_ACU_5= SheetBanco.getRange(Linha, 483).getValue();
    var RG_ACU_5= SheetBanco.getRange(Linha, 484).getValue();
    var CPF_ACU_5= SheetBanco.getRange(Linha, 485).getValue();
    var SX_ACU_5= SheetBanco.getRange(Linha, 486).getValue();
    var DN_ACU_5= SheetBanco.getRange(Linha, 487).getValue();
    var ORIEN_SX_ACU_5= SheetBanco.getRange(Linha, 491).getValue();
    var COR_ACU_5= SheetBanco.getRange(Linha, 492).getValue();
    var EC_ACU_5= SheetBanco.getRange(Linha, 493).getValue();
    var UE_ACU_5= SheetBanco.getRange(Linha, 494).getValue();
    var FIL_ACU_5= SheetBanco.getRange(Linha, 495).getValue();
    var NAT_ACU_5= SheetBanco.getRange(Linha, 496).getValue();
    var NAC_ACU_5= SheetBanco.getRange(Linha, 497).getValue();
    var COND_FIS_ACU_5= SheetBanco.getRange(Linha, 498).getValue();
    var ALCUNHA_ACU_5= SheetBanco.getRange(Linha, 499).getValue();
    var TAT_ACU_5= SheetBanco.getRange(Linha, 500).getValue();
    var ESC_ACU_5= SheetBanco.getRange(Linha, 501).getValue();
    //var nome_var = SheetBanco.getRange(Linha, col).getValue();
    var END_RES_ACU_5= SheetBanco.getRange(Linha, 502).getValue();
    var MUN_ACU_5= SheetBanco.getRange(Linha, 503).getValue();
    var BAIRRO_ACU_5= SheetBanco.getRange(Linha, 504).getValue();
    var PROF_ACU_5= SheetBanco.getRange(Linha, 505).getValue();
    var END_PROF_ACU_5= SheetBanco.getRange(Linha, 506).getValue();
    var MUN_PROF_ACU_5= SheetBanco.getRange(Linha, 507).getValue();
    var BAIRRO_PROF_ACU_5= SheetBanco.getRange(Linha, 508).getValue();
    var N_PAI_ACU_5= SheetBanco.getRange(Linha, 509).getValue();
    var ESC_PAI_ACU_5= SheetBanco.getRange(Linha, 510).getValue();
    var HIS_AC_PAI_ACU_5= SheetBanco.getRange(Linha, 511).getValue();
    var HIS_AC_PAI_VDOM_ACU_5= SheetBanco.getRange(Linha, 512).getValue();
    //var nome_var = SheetBanco.getRange(Linha, col).getValue();
    var ESC_MAE_ACU_5= SheetBanco.getRange(Linha, 513).getValue();
    var N_MAE_ACU_5= SheetBanco.getRange(Linha, 514).getValue();
    var HIS_AC_MAE_ACU_5= SheetBanco.getRange(Linha, 515).getValue();
    var VIOL_DOM_MAE_ACU_5= SheetBanco.getRange(Linha, 516).getValue();
    //var nome_var = SheetBanco.getRange(Linha, col).getValue();
    var MP_ACU_5= SheetBanco.getRange(Linha, 517).getValue();
    var MI_ACU_5= SheetBanco.getRange(Linha, 518).getValue();
    var HIS_MI_ACU_5= SheetBanco.getRange(Linha, 519).getValue();
    var ENVTRAF_ACU_5= SheetBanco.getRange(Linha, 520).getValue();
    var FOR_ACU_5= SheetBanco.getRange(Linha, 521).getValue();
    var ANTEPOL_ACU_5= SheetBanco.getRange(Linha, 522).getValue();
    var ANTECRIM_ACU_5= SheetBanco.getRange(Linha, 523).getValue();
    //var nome_var = SheetBanco.getRange(Linha, col).getValue();
    //var nome_var = SheetBanco.getRange(Linha, col).getValue();
    var ANTEC_TRANS_ACU_5= SheetBanco.getRange(Linha, 524).getValue();
    var ANTEC_DIR_HAB_ACU_5= SheetBanco.getRange(Linha, 525).getValue();
    var ANTEC_FUG_RECAP_ACU_5= SheetBanco.getRange(Linha, 526).getValue();
    var ANTEC_ENT_POS_ACU_5= SheetBanco.getRange(Linha, 527).getValue();
    var ANTEC_ENT_TRAF_ACU_5= SheetBanco.getRange(Linha, 528).getValue();
    var ANTEC_PORT_ARMAS_ACU_5= SheetBanco.getRange(Linha, 529).getValue();
    var ANTEC_RECEP_ACU_5= SheetBanco.getRange(Linha, 530).getValue();
    var ANTEC_FURTO_ACU_5= SheetBanco.getRange(Linha, 531).getValue();
    var ANTEC_ROUBO_PED_ACU_5= SheetBanco.getRange(Linha, 532).getValue();
    var ANTEC_ROUBO_BANC_ACU_5= SheetBanco.getRange(Linha, 533).getValue();
    var ANTEC_ROUBO_TRANSP_PUB_ACU_5= SheetBanco.getRange(Linha, 534).getValue();
    var ANTEC_ROUBO_TRANSP_IND_ACU_5= SheetBanco.getRange(Linha, 535).getValue();
    var ANTEC_ROUBO_TRANSP_IND_APP_ACU_5= SheetBanco.getRange(Linha, 536).getValue();
    var ANTEC_ROUBO_TRANSP_VEIC_ACU_5= SheetBanco.getRange(Linha, 537).getValue();
    var ANTEC_ROUBO_RESID_ACU_5= SheetBanco.getRange(Linha, 538).getValue();
    var ANTEC_ROUBO_OUT_ACU_5= SheetBanco.getRange(Linha, 539).getValue();
    var ANTEC_LATROC_ACU_5= SheetBanco.getRange(Linha, 540).getValue();
    var ANTEC_AME_ACU_5= SheetBanco.getRange(Linha, 541).getValue();
    var ANTEC_LES_CORP_ACU_5= SheetBanco.getRange(Linha, 542).getValue();
    var ANTEC_LES_CORP_MOR_ACU_5= SheetBanco.getRange(Linha, 543).getValue();
    var ANTEC_HOM_ACU_5= SheetBanco.getRange(Linha, 544).getValue();
    var ANTEC_MP_ACU_5= SheetBanco.getRange(Linha, 545).getValue();
    var ANTEC_CRIM_SEX_ACU_5= SheetBanco.getRange(Linha, 546).getValue();
    var ANTEC_ESTUP_VUL_ACU_5= SheetBanco.getRange(Linha, 547).getValue();
    var ANTEC_OUT_ACU_5= SheetBanco.getRange(Linha, 548).getValue();
    var HIST_VIT_TRANS_ACU_5= SheetBanco.getRange(Linha, 549).getValue();
    var HIST_VIT_FURTO_ACU_5= SheetBanco.getRange(Linha, 550).getValue();
    var HIST_VIT_ROUBO_PED_ACU_5= SheetBanco.getRange(Linha, 551).getValue();
    var HIST_VIT_ROUBO_BANC_ACU_5= SheetBanco.getRange(Linha, 552).getValue();
    var HIST_VIT_ROUBO_TRANSP_PUB_ACU_5= SheetBanco.getRange(Linha, 553).getValue();
    var HIST_VIT_ROUBO_TRANSP_IND_ACU_5= SheetBanco.getRange(Linha, 554).getValue();
    var HIST_VIT_ROUBO_TRANSP_IND_APP_ACU_5= SheetBanco.getRange(Linha, 555).getValue();
    var HIST_VIT_ROUBO_TRANSP_VEIC_ACU_5= SheetBanco.getRange(Linha, 556).getValue();
    var HIST_VIT_ROUBO_RESID_ACU_5= SheetBanco.getRange(Linha, 557).getValue();
    var HIST_VIT_ROUBO_OUT_ACU_5= SheetBanco.getRange(Linha, 558).getValue();
    var HIST_VIT_LATROC_ACU_5= SheetBanco.getRange(Linha, 559).getValue();
    var HIST_VIT_AME_ACU_5= SheetBanco.getRange(Linha, 560).getValue();
    var HIST_VIT_LES_CORP_ACU_5= SheetBanco.getRange(Linha, 561).getValue();
    var HIST_VIT_HOM_ACU_5= SheetBanco.getRange(Linha, 562).getValue();
    var HIST_VIT_MP_ACU_5= SheetBanco.getRange(Linha, 563).getValue();
    var HIST_VIT_CRIM_SEX_ACU_5= SheetBanco.getRange(Linha, 564).getValue();
    var HIST_VIT_ESTUP_VUL_ACU_5= SheetBanco.getRange(Linha, 565).getValue();
    var HIST_VIT_OUT_ACU_5= SheetBanco.getRange(Linha, 566).getValue();
    //var nome_var = SheetBanco.getRange(Linha, col).getValue();
    //var nome_var = SheetBanco.getRange(Linha, col).getValue();
    //var nome_var = SheetBanco.getRange(Linha, col).getValue();
    var N_ACU_6= SheetBanco.getRange(Linha, 567).getValue();
    var RG_ACU_6= SheetBanco.getRange(Linha, 568).getValue();
    var CPF_ACU_6= SheetBanco.getRange(Linha, 569).getValue();
    var SX_ACU_6= SheetBanco.getRange(Linha, 570).getValue();
    var DN_ACU_6= SheetBanco.getRange(Linha, 571).getValue();
    var ORIEN_SX_ACU_6= SheetBanco.getRange(Linha, 575).getValue();
    var COR_ACU_6= SheetBanco.getRange(Linha, 576).getValue();
    var EC_ACU_6= SheetBanco.getRange(Linha, 577).getValue();
    var UE_ACU_6= SheetBanco.getRange(Linha, 578).getValue();
    var FIL_ACU_6= SheetBanco.getRange(Linha, 579).getValue();
    var NAT_ACU_6= SheetBanco.getRange(Linha, 580).getValue();
    var NAC_ACU_6= SheetBanco.getRange(Linha, 581).getValue();
    var COND_FIS_ACU_6= SheetBanco.getRange(Linha, 582).getValue();
    var ALCUNHA_ACU_6= SheetBanco.getRange(Linha, 583).getValue();
    var TAT_ACU_6= SheetBanco.getRange(Linha, 584).getValue();
    var ESC_ACU_6= SheetBanco.getRange(Linha, 585).getValue();
    //var nome_var = SheetBanco.getRange(Linha, col).getValue();
    var END_RES_ACU_6= SheetBanco.getRange(Linha, 586).getValue();
    var MUN_ACU_6= SheetBanco.getRange(Linha, 587).getValue();
    var BAIRRO_ACU_6= SheetBanco.getRange(Linha, 588).getValue();
    var PROF_ACU_6= SheetBanco.getRange(Linha, 589).getValue();
    var END_PROF_ACU_6= SheetBanco.getRange(Linha, 590).getValue();
    var MUN_PROF_ACU_6= SheetBanco.getRange(Linha, 591).getValue();
    var BAIRRO_PROF_ACU_6= SheetBanco.getRange(Linha, 592).getValue();
    var N_PAI_ACU_6= SheetBanco.getRange(Linha, 593).getValue();
    var ESC_PAI_ACU_6= SheetBanco.getRange(Linha, 594).getValue();
    var HIS_AC_PAI_ACU_6= SheetBanco.getRange(Linha, 595).getValue();
    var HIS_AC_PAI_VDOM_ACU_6= SheetBanco.getRange(Linha, 596).getValue();
    //var nome_var = SheetBanco.getRange(Linha, col).getValue();
    var ESC_MAE_ACU_6= SheetBanco.getRange(Linha, 597).getValue();
    var N_MAE_ACU_6= SheetBanco.getRange(Linha, 598).getValue();
    var HIS_AC_MAE_ACU_6= SheetBanco.getRange(Linha, 599).getValue();
    var VIOL_DOM_MAE_ACU_6= SheetBanco.getRange(Linha, 600).getValue();
    //var nome_var = SheetBanco.getRange(Linha, col).getValue();
    var MP_ACU_6= SheetBanco.getRange(Linha, 601).getValue();
    var MI_ACU_6= SheetBanco.getRange(Linha, 602).getValue();
    var HIS_MI_ACU_6= SheetBanco.getRange(Linha, 603).getValue();
    var ENVTRAF_ACU_6= SheetBanco.getRange(Linha, 604).getValue();
    var FOR_ACU_6= SheetBanco.getRange(Linha, 605).getValue();
    var ANTEPOL_ACU_6= SheetBanco.getRange(Linha, 606).getValue();
    var ANTECRIM_ACU_6= SheetBanco.getRange(Linha, 607).getValue();
    //var nome_var = SheetBanco.getRange(Linha, col).getValue();
    //var nome_var = SheetBanco.getRange(Linha, col).getValue();
    var ANTEC_TRANS_ACU_6= SheetBanco.getRange(Linha, 608).getValue();
    var ANTEC_DIR_HAB_ACU_6= SheetBanco.getRange(Linha, 609).getValue();
    var ANTEC_FUG_RECAP_ACU_6= SheetBanco.getRange(Linha, 610).getValue();
    var ANTEC_ENT_POS_ACU_6= SheetBanco.getRange(Linha, 611).getValue();
    var ANTEC_ENT_TRAF_ACU_6= SheetBanco.getRange(Linha, 612).getValue();
    var ANTEC_PORT_ARMAS_ACU_6= SheetBanco.getRange(Linha, 613).getValue();
    var ANTEC_RECEP_ACU_6= SheetBanco.getRange(Linha, 614).getValue();
    var ANTEC_FURTO_ACU_6= SheetBanco.getRange(Linha, 615).getValue();
    var ANTEC_ROUBO_PED_ACU_6= SheetBanco.getRange(Linha, 616).getValue();
    var ANTEC_ROUBO_BANC_ACU_6= SheetBanco.getRange(Linha, 617).getValue();
    var ANTEC_ROUBO_TRANSP_PUB_ACU_6= SheetBanco.getRange(Linha, 618).getValue();
    var ANTEC_ROUBO_TRANSP_IND_ACU_6= SheetBanco.getRange(Linha, 619).getValue();
    var ANTEC_ROUBO_TRANSP_IND_APP_ACU_6= SheetBanco.getRange(Linha, 620).getValue();
    var ANTEC_ROUBO_TRANSP_VEIC_ACU_6= SheetBanco.getRange(Linha, 621).getValue();
    var ANTEC_ROUBO_RESID_ACU_6= SheetBanco.getRange(Linha, 622).getValue();
    var ANTEC_ROUBO_OUT_ACU_6= SheetBanco.getRange(Linha, 623).getValue();
    var ANTEC_LATROC_ACU_6= SheetBanco.getRange(Linha, 624).getValue();
    var ANTEC_AME_ACU_6= SheetBanco.getRange(Linha, 625).getValue();
    var ANTEC_LES_CORP_ACU_6= SheetBanco.getRange(Linha, 626).getValue();
    var ANTEC_LES_CORP_MOR_ACU_6= SheetBanco.getRange(Linha, 627).getValue();
    var ANTEC_HOM_ACU_6= SheetBanco.getRange(Linha, 628).getValue();
    var ANTEC_MP_ACU_6= SheetBanco.getRange(Linha, 629).getValue();
    var ANTEC_CRIM_SEX_ACU_6= SheetBanco.getRange(Linha, 630).getValue();
    var ANTEC_ESTUP_VUL_ACU_6= SheetBanco.getRange(Linha, 631).getValue();
    var ANTEC_OUT_ACU_6= SheetBanco.getRange(Linha, 632).getValue();
    var HIST_VIT_TRANS_ACU_6= SheetBanco.getRange(Linha, 633).getValue();
    var HIST_VIT_FURTO_ACU_6= SheetBanco.getRange(Linha, 634).getValue();
    var HIST_VIT_ROUBO_PED_ACU_6= SheetBanco.getRange(Linha, 635).getValue();
    var HIST_VIT_ROUBO_BANC_ACU_6= SheetBanco.getRange(Linha, 636).getValue();
    var HIST_VIT_ROUBO_TRANSP_PUB_ACU_6= SheetBanco.getRange(Linha, 637).getValue();
    var HIST_VIT_ROUBO_TRANSP_IND_ACU_6= SheetBanco.getRange(Linha, 638).getValue();
    var HIST_VIT_ROUBO_TRANSP_IND_APP_ACU_6= SheetBanco.getRange(Linha, 639).getValue();
    var HIST_VIT_ROUBO_TRANSP_VEIC_ACU_6= SheetBanco.getRange(Linha, 640).getValue();
    var HIST_VIT_ROUBO_RESID_ACU_6= SheetBanco.getRange(Linha, 641).getValue();
    var HIST_VIT_ROUBO_OUT_ACU_6= SheetBanco.getRange(Linha, 642).getValue();
    var HIST_VIT_LATROC_ACU_6= SheetBanco.getRange(Linha, 643).getValue();
    var HIST_VIT_AME_ACU_6= SheetBanco.getRange(Linha, 644).getValue();
    var HIST_VIT_LES_CORP_ACU_6= SheetBanco.getRange(Linha, 645).getValue();
    var HIST_VIT_HOM_ACU_6= SheetBanco.getRange(Linha, 646).getValue();
    var HIST_VIT_MP_ACU_6= SheetBanco.getRange(Linha, 647).getValue();
    var HIST_VIT_CRIM_SEX_ACU_6= SheetBanco.getRange(Linha, 648).getValue();
    var HIST_VIT_ESTUP_VUL_ACU_6= SheetBanco.getRange(Linha, 649).getValue();
    var HIST_VIT_OUT_ACU_6= SheetBanco.getRange(Linha, 650).getValue();
    //var nome_var = SheetBanco.getRange(Linha, col).getValue();
    //var nome_var = SheetBanco.getRange(Linha, col).getValue();
    //var nome_var = SheetBanco.getRange(Linha, col).getValue();
    
    Entrada.getRange('B2').activate();
    Entrada.getCurrentCell().setValue(NC);
    
    Entrada.getRange('B3').activate();
    Entrada.getCurrentCell().setValue(NUM_FATO);
    
    Entrada.getRange('B4').activate();
    Entrada.getCurrentCell().setValue(NUM_OCOR);
    
    Entrada.getRange('B5').activate();
    Entrada.getCurrentCell().setValue(TIP_FATO);
    
    Entrada.getRange('B6').activate();
    Entrada.getCurrentCell().setValue(NUM_ORG_REGIST);
    
    Entrada.getRange('B7').activate();
    Entrada.getCurrentCell().setValue(N_ORG_REGIST);
    
    Entrada.getRange('B8').activate();
    Entrada.getCurrentCell().setValue(ORIGEM_COM);
    
    Entrada.getRange('B9').activate();
    Entrada.getCurrentCell().setValue(DATA_COM);
    
    Entrada.getRange('B10').activate();
    Entrada.getCurrentCell().setValue(HORA_COM);
    
    Entrada.getRange('B11').activate();
    Entrada.getCurrentCell().setValue(ANO_REGIST);
    
    Entrada.getRange('B12').activate();
    Entrada.getCurrentCell().setValue(FLAGRANTE);
    
    Entrada.getRange('B13').activate();
    Entrada.getCurrentCell().setValue(TENTATIVA);
    
    Entrada.getRange('B14').activate();
    Entrada.getCurrentCell().setValue(COMUM);
    
    Entrada.getRange('B15').activate();
    Entrada.getCurrentCell().setValue(COND);
    
    Entrada.getRange('B16').activate();
    Entrada.getCurrentCell().setValue(LAUDO);
    
    //Entrada.getRange('B17').activate();
    //Entrada.getCurrentCell().setValue();
    
    Entrada.getRange('D2').activate();
    Entrada.getCurrentCell().setValue(NUM_AUT);
    
    Entrada.getRange('D3').activate();
    Entrada.getCurrentCell().setValue(NUM_VIT);
    
    Entrada.getRange('D4').activate();
    Entrada.getCurrentCell().setValue(NUM_VIT_MORTAS);
    
    Entrada.getRange('D5').activate();
    Entrada.getCurrentCell().setValue(NUMTEST);
    
    Entrada.getRange('D6').activate();
    Entrada.getCurrentCell().setValue(TIPO_TEST);
    
    Entrada.getRange('D8').activate();
    Entrada.getCurrentCell().setValue(DATA_FATO);
    
    Entrada.getRange('D11').activate();
    Entrada.getCurrentCell().setValue(H_FATO);
    
    Entrada.getRange('D13').activate();
    Entrada.getCurrentCell().setValue(MUN);
    
    Entrada.getRange('F8').activate(); 
    Entrada.getCurrentCell().setValue(A_MUN);
    
    Entrada.getRange('F9').activate(); 
    Entrada.getCurrentCell().setValue(LOG);
    
    Entrada.getRange('F10').activate(); 
    Entrada.getCurrentCell().setValue(NUM);
    
    Entrada.getRange('F11').activate(); 
    Entrada.getCurrentCell().setValue(COMP);
    
    Entrada.getRange('F12').activate(); 
    Entrada.getCurrentCell().setValue(CEP);
    
    Entrada.getRange('F13').activate(); 
    Entrada.getCurrentCell().setValue(BAIRRO);
    
    Entrada.getRange('F14').activate(); 
    Entrada.getCurrentCell().setValue(LAT);
    
    Entrada.getRange('F15').activate(); 
    Entrada.getCurrentCell().setValue(LNG);
    
    Entrada.getRange('F16').activate(); 
    Entrada.getCurrentCell().setValue(P_REF);
    
    Entrada.getRange('F17').activate(); 
    Entrada.getCurrentCell().setValue(TIPO_LOCAL);
    
    Entrada.getRange('H2').activate(); 
    Entrada.getCurrentCell().setValue(M_UTIL_AUT);
    
    Entrada.getRange('H3').activate(); 
    Entrada.getCurrentCell().setValue(RECUR);
    
    Entrada.getRange('H4').activate(); 
    Entrada.getCurrentCell().setValue(NUM_AGRE);
    
    Entrada.getRange('H5').activate(); 
    Entrada.getCurrentCell().setValue(M_AUTOR);
    
    Entrada.getRange('H6').activate(); 
    Entrada.getCurrentCell().setValue(OBJ);
    
    Entrada.getRange('H7').activate(); 
    Entrada.getCurrentCell().setValue(ADTNT);
    
    Entrada.getRange('H8').activate(); 
    Entrada.getCurrentCell().setValue(MOT);
    
    Entrada.getRange('H9').activate(); 
    Entrada.getCurrentCell().setValue(CADAVER);
    
    Entrada.getRange('H10').activate(); 
    Entrada.getCurrentCell().setValue(R_VIT_AUT);
    
    Entrada.getRange('H11').activate(); 
    Entrada.getCurrentCell().setValue(MID);
    
    Entrada.getRange('H12').activate(); 
    Entrada.getCurrentCell().setValue(INFORMANTE);
    
    Entrada.getRange('H13').activate(); 
    Entrada.getCurrentCell().setValue(N_INFORMANTE);
    
    Entrada.getRange('H14').activate(); 
    Entrada.getCurrentCell().setValue(HIST);
    
    Entrada.getRange('H16').activate(); 
    Entrada.getCurrentCell().setValue(OBS);
    
    Entrada.getRange('B19').activate(); 
    Entrada.getCurrentCell().setValue(N_VIT);
    
    Entrada.getRange('B20').activate(); 
    Entrada.getCurrentCell().setValue(RG_VIT);
    
    Entrada.getRange('B21').activate(); 
    Entrada.getCurrentCell().setValue(CPF_VIT);
    
    Entrada.getRange('B22').activate(); 
    Entrada.getCurrentCell().setValue(SX_VIT);
    
    Entrada.getRange('B23').activate(); 
    Entrada.getCurrentCell().setValue(DN_VIT);
    
    Entrada.getRange('B27').activate(); 
    Entrada.getCurrentCell().setValue(ORIEN_SX_VIT);
    
    Entrada.getRange('B28').activate(); 
    Entrada.getCurrentCell().setValue(COR_VIT);
    
    Entrada.getRange('B29').activate(); 
    Entrada.getCurrentCell().setValue(EC_VIT);
    
    Entrada.getRange('B30').activate(); 
    Entrada.getCurrentCell().setValue(UE_VIT);
    
    Entrada.getRange('B31').activate(); 
    Entrada.getCurrentCell().setValue(FIL_VIT);
    
    Entrada.getRange('B32').activate(); 
    Entrada.getCurrentCell().setValue(NAT_VIT);
    
    Entrada.getRange('B33').activate(); 
    Entrada.getCurrentCell().setValue(NAC_VIT);
    
    Entrada.getRange('B34').activate(); 
    Entrada.getCurrentCell().setValue(COND_FIS_VIT);
    
    Entrada.getRange('B35').activate(); 
    Entrada.getCurrentCell().setValue(ALCUNHA_VIT);
    
    Entrada.getRange('B36').activate(); 
    Entrada.getCurrentCell().setValue(TAT_VIT);
    
    Entrada.getRange('B37').activate(); 
    Entrada.getCurrentCell().setValue(ESC_VIT);
    
    //Entrada.getRange('B38').activate(); 
    //Entrada.getCurrentCell().setValue();
    
    Entrada.getRange('B39').activate(); 
    Entrada.getCurrentCell().setValue(END_RES_VIT);
    
    Entrada.getRange('B40').activate(); 
    Entrada.getCurrentCell().setValue(MUN_VIT);
    
    Entrada.getRange('B41').activate(); 
    Entrada.getCurrentCell().setValue(BAIRRO_VIT);
    
    Entrada.getRange('D19').activate(); 
    Entrada.getCurrentCell().setValue(PROF_VIT);
    
    Entrada.getRange('D20').activate(); 
    Entrada.getCurrentCell().setValue(END_PROF_VIT);
    
    Entrada.getRange('D21').activate(); 
    Entrada.getCurrentCell().setValue(MUN_PROF_VIT);
    
    Entrada.getRange('D22').activate(); 
    Entrada.getCurrentCell().setValue(BAIRRO_PROF_VIT);
    
    Entrada.getRange('D23').activate(); 
    Entrada.getCurrentCell().setValue(N_PAI_VIT);
    
    Entrada.getRange('D24').activate(); 
    Entrada.getCurrentCell().setValue(ESC_PAI_VIT);
    
    Entrada.getRange('D25').activate(); 
    Entrada.getCurrentCell().setValue(HIS_AC_PAI_VIT);
    
    Entrada.getRange('D26').activate(); 
    Entrada.getCurrentCell().setValue(HIS_AC_PAI_VDOM_VIT);
    
    //Entrada.getRange('D27').activate(); 
    //Entrada.getCurrentCell().setValue();
    
    Entrada.getRange('D28').activate(); 
    Entrada.getCurrentCell().setValue(N_MAE_VIT);
    
    Entrada.getRange('D29').activate(); 
    Entrada.getCurrentCell().setValue(ESC_MAEVIT);
    
    Entrada.getRange('D30').activate(); 
    Entrada.getCurrentCell().setValue(HIS_AC_MAE_VIT);
    
    Entrada.getRange('D31').activate(); 
    Entrada.getCurrentCell().setValue(VIOL_DOM_MAE_VIT);
    
    //Entrada.getRange('D32').activate(); 
    //Entrada.getCurrentCell().setValue();
    
    Entrada.getRange('D33').activate(); 
    Entrada.getCurrentCell().setValue(MP_VIT);
    
    Entrada.getRange('D34').activate(); 
    Entrada.getCurrentCell().setValue(MI_VIT);
    
    Entrada.getRange('D35').activate(); 
    Entrada.getCurrentCell().setValue(HIS_MI_VIT);
    
    Entrada.getRange('D36').activate(); 
    Entrada.getCurrentCell().setValue(ENVTRAF_VIT);
    
    Entrada.getRange('D37').activate(); 
    Entrada.getCurrentCell().setValue(FOR_VIT);
    
    Entrada.getRange('D38').activate(); 
    Entrada.getCurrentCell().setValue(ANTEPOL_VIT);
    
    Entrada.getRange('D39').activate(); 
    Entrada.getCurrentCell().setValue(ANTECRIM_VIT);
    
    //Entrada.getRange('D40').activate(); 
    //Entrada.getCurrentCell().setValue();
    
    //Entrada.getRange('D41').activate(); 
    //Entrada.getCurrentCell().setValue();
    
    Entrada.getRange('F19').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_TRANS_VIT);
    
    Entrada.getRange('F20').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_DIR_HAB_VIT);
    
    Entrada.getRange('F21').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_FUG_RECAP_VIT);
    
    Entrada.getRange('F22').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ENT_POS_VIT);
    
    Entrada.getRange('F23').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ENT_TRAF_VIT);
    
    Entrada.getRange('F24').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_PORT_ARMAS_VIT);
    
    Entrada.getRange('F25').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_RECEP_VIT);
    
    Entrada.getRange('F26').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_FURTO_VIT);
    
    Entrada.getRange('F27').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ROUBO_PED_VIT);
    
    Entrada.getRange('F28').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ROUBO_BANC_VIT);
    
    Entrada.getRange('F29').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ROUBO_TRANSP_PUB_VIT);
    
    Entrada.getRange('F30').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ROUBO_TRANSP_IND_VIT);
    
    Entrada.getRange('F31').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ROUBO_TRANSP_IND_APP_VIT);
    
    Entrada.getRange('F32').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ROUBO_TRANSP_VEIC_VIT);
    
    Entrada.getRange('F33').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ROUBO_RESID_VIT);
    
    Entrada.getRange('F34').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ROUBO_OUT_VIT);
    
    Entrada.getRange('F35').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_LATROC_VIT);
    
    Entrada.getRange('F36').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_AME_VIT);
    
    Entrada.getRange('F37').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_LES_CORP_VIT);
    
    Entrada.getRange('F38').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_LES_CORP_MOR_VIT);
    
    Entrada.getRange('F39').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_HOM_VIT);
    
    Entrada.getRange('F40').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_MP_VIT);
    
    Entrada.getRange('F41').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_CRIM_SEX_VIT);
    
    Entrada.getRange('H19').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ESTUP_VUL_VIT);
    
    Entrada.getRange('H20').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_OUT_VIT);
    
    Entrada.getRange('H21').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_TRANS_VIT);
    
    Entrada.getRange('H22').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_FURTO_VIT);
    
    Entrada.getRange('H23').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_ROUBO_PED_VIT);
    
    Entrada.getRange('H24').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_ROUBO_BANC_VIT);
    
    Entrada.getRange('H25').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_ROUBO_TRANSP_PUB_VIT);
    
    Entrada.getRange('H26').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_ROUBO_TRANSP_IND_VIT);
    
    Entrada.getRange('H27').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_ROUBO_TRANSP_IND_APP_VIT);
    
    Entrada.getRange('H28').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_ROUBO_TRANSP_VEIC_VIT);
    
    Entrada.getRange('H29').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_ROUBO_RESID_VIT);
    
    Entrada.getRange('H30').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_ROUBO_OUT_VIT);
    
    Entrada.getRange('H31').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_LATROC_VIT);
    
    Entrada.getRange('H32').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_AME_VIT);
    
    Entrada.getRange('H33').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_LES_CORP_VIT);
    
    Entrada.getRange('H34').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_HOM_VIT);
    
    Entrada.getRange('H35').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_MP_VIT);
    
    Entrada.getRange('H36').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_CRIM_SEX_VIT);
    
    Entrada.getRange('H37').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_ESTUP_VUL_VIT);
    
    Entrada.getRange('H38').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_OUT_VIT);
    
    //Entrada.getRange('H39').activate(); 
    //Entrada.getCurrentCell().setValue();
    
    //Entrada.getRange('H40').activate(); 
    //Entrada.getCurrentCell().setValue();
    
    //Entrada.getRange('H41').activate(); 
    //Entrada.getCurrentCell().setValue();
    
    Entrada.getRange('B43').activate(); 
    Entrada.getCurrentCell().setValue(N_ACU_1);
    
    Entrada.getRange('B44').activate(); 
    Entrada.getCurrentCell().setValue(RG_ACU_1);
    
    Entrada.getRange('B45').activate(); 
    Entrada.getCurrentCell().setValue(CPF_ACU_1);
    
    Entrada.getRange('B46').activate(); 
    Entrada.getCurrentCell().setValue(SX_ACU_1);
    
    Entrada.getRange('B47').activate(); 
    Entrada.getCurrentCell().setValue(DN_ACU_1);
    
    Entrada.getRange('B51').activate(); 
    Entrada.getCurrentCell().setValue(ORIEN_SX_ACU_1);
    
    Entrada.getRange('B52').activate(); 
    Entrada.getCurrentCell().setValue(COR_ACU_1);
    
    Entrada.getRange('B53').activate(); 
    Entrada.getCurrentCell().setValue(EC_ACU_1);
    
    Entrada.getRange('B54').activate(); 
    Entrada.getCurrentCell().setValue(UE_ACU_1);
    
    Entrada.getRange('B55').activate(); 
    Entrada.getCurrentCell().setValue(FIL_ACU_1);
    
    Entrada.getRange('B56').activate(); 
    Entrada.getCurrentCell().setValue(NAT_ACU_1);
    
    Entrada.getRange('B57').activate(); 
    Entrada.getCurrentCell().setValue(NAC_ACU_1);
    
    Entrada.getRange('B58').activate(); 
    Entrada.getCurrentCell().setValue(COND_FIS_ACU_1);
    
    Entrada.getRange('B59').activate(); 
    Entrada.getCurrentCell().setValue(ALCUNHA_ACU_1);
    
    Entrada.getRange('B60').activate(); 
    Entrada.getCurrentCell().setValue(TAT_ACU_1);
    
    Entrada.getRange('B61').activate(); 
    Entrada.getCurrentCell().setValue(ESC_ACU_1);
    
    //Entrada.getRange('B62').activate(); 
    //Entrada.getCurrentCell().setValue();
    
    Entrada.getRange('B63').activate(); 
    Entrada.getCurrentCell().setValue(END_RES_ACU_1);
    
    Entrada.getRange('B64').activate(); 
    Entrada.getCurrentCell().setValue(MUN_ACU_1);
    
    Entrada.getRange('B65').activate(); 
    Entrada.getCurrentCell().setValue(BAIRRO_ACU_1);
    
    Entrada.getRange('B66').activate(); 
    Entrada.getCurrentCell().setValue(PROF_ACU_1);
    
    Entrada.getRange('B67').activate(); 
    Entrada.getCurrentCell().setValue(END_PROF_ACU_1);
    
    Entrada.getRange('B68').activate(); 
    Entrada.getCurrentCell().setValue(MUN_PROF_ACU_1);
    
    Entrada.getRange('B69').activate(); 
    Entrada.getCurrentCell().setValue(BAIRRO_PROF_ACU_1);
    
    Entrada.getRange('B70').activate(); 
    Entrada.getCurrentCell().setValue(N_PAI_ACU_1);
    
    Entrada.getRange('B71').activate(); 
    Entrada.getCurrentCell().setValue(ESC_PAI_ACU_1);
    
    Entrada.getRange('B72').activate(); 
    Entrada.getCurrentCell().setValue(HIS_AC_PAI_ACU_1);
    
    Entrada.getRange('B73').activate(); 
    Entrada.getCurrentCell().setValue(HIS_AC_PAI_VDOM_ACU_1);
    
    //Entrada.getRange('B74').activate(); 
    //Entrada.getCurrentCell().setValue();
    
    Entrada.getRange('B75').activate(); 
    Entrada.getCurrentCell().setValue(ESC_MAE_ACU_1);
    
    Entrada.getRange('B76').activate(); 
    Entrada.getCurrentCell().setValue(N_MAE_ACU_1);
    
    Entrada.getRange('B77').activate(); 
    Entrada.getCurrentCell().setValue(HIS_AC_MAE_ACU_1);
    
    Entrada.getRange('B78').activate(); 
    Entrada.getCurrentCell().setValue(VIOL_DOM_MAE_ACU_1);
    
    //Entrada.getRange('B79').activate(); 
    //Entrada.getCurrentCell().setValue();
    
    Entrada.getRange('B80').activate(); 
    Entrada.getCurrentCell().setValue(MP_ACU_1);
    
    Entrada.getRange('B81').activate(); 
    Entrada.getCurrentCell().setValue(MI_ACU_1);
    
    Entrada.getRange('B82').activate(); 
    Entrada.getCurrentCell().setValue(HIS_MI_ACU_1);
    
    Entrada.getRange('B83').activate(); 
    Entrada.getCurrentCell().setValue(ENVTRAF_ACU_1);
    
    Entrada.getRange('B84').activate(); 
    Entrada.getCurrentCell().setValue(FOR_ACU_1);
    
    Entrada.getRange('B85').activate(); 
    Entrada.getCurrentCell().setValue(ANTEPOL_ACU_1);
    
    Entrada.getRange('B86').activate(); 
    Entrada.getCurrentCell().setValue(ANTECRIM_ACU_1);
    
    //Entrada.getRange('B87').activate(); 
    //Entrada.getCurrentCell().setValue();
    
    //Entrada.getRange('B88').activate(); 
    //Entrada.getCurrentCell().setValue();
    
    Entrada.getRange('D43').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_TRANS_ACU_1);
    
    Entrada.getRange('D44').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_DIR_HAB_ACU_1);
    
    Entrada.getRange('D45').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_FUG_RECAP_ACU_1);
    
    Entrada.getRange('D46').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ENT_POS_ACU_1);
    
    Entrada.getRange('D47').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ENT_TRAF_ACU_1);
    
    Entrada.getRange('D48').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_PORT_ARMAS_ACU_1);
    
    Entrada.getRange('D49').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_RECEP_ACU_1);
    
    Entrada.getRange('D50').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_FURTO_ACU_1);
    
    Entrada.getRange('D51').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ROUBO_PED_ACU_1);
    
    Entrada.getRange('D52').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ROUBO_BANC_ACU_1);
    
    Entrada.getRange('D53').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ROUBO_TRANSP_PUB_ACU_1);
    
    Entrada.getRange('D54').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ROUBO_TRANSP_IND_ACU_1);
    
    Entrada.getRange('D55').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ROUBO_TRANSP_IND_APP_ACU_1);
    
    Entrada.getRange('D56').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ROUBO_TRANSP_VEIC_ACU_1);
    
    Entrada.getRange('D57').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ROUBO_RESID_ACU_1);
    
    Entrada.getRange('D58').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ROUBO_OUT_ACU_1);
    
    Entrada.getRange('D59').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_LATROC_ACU_1);
    
    Entrada.getRange('D60').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_AME_ACU_1);
    
    Entrada.getRange('D61').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_LES_CORP_ACU_1);
    
    Entrada.getRange('D62').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_LES_CORP_MOR_ACU_1);
    
    Entrada.getRange('D63').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_HOM_ACU_1);
    
    Entrada.getRange('D64').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_MP_ACU_1);
    
    Entrada.getRange('D65').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_CRIM_SEX_ACU_1);
    
    Entrada.getRange('D66').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ESTUP_VUL_ACU_1);
    
    Entrada.getRange('D67').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_OUT_ACU_1);
    
    Entrada.getRange('D68').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_TRANS_ACU_1);
    
    Entrada.getRange('D69').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_FURTO_ACU_1);
    
    Entrada.getRange('D70').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_ROUBO_PED_ACU_1);
    
    Entrada.getRange('D71').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_ROUBO_BANC_ACU_1);
    
    Entrada.getRange('D72').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_ROUBO_TRANSP_PUB_ACU_1);
    
    Entrada.getRange('D73').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_ROUBO_TRANSP_IND_ACU_1);
    
    Entrada.getRange('D74').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_ROUBO_TRANSP_IND_APP_ACU_1);
    
    Entrada.getRange('D75').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_ROUBO_TRANSP_VEIC_ACU_1);
    
    Entrada.getRange('D76').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_ROUBO_RESID_ACU_1);
    
    Entrada.getRange('D77').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_ROUBO_OUT_ACU_1);
    
    Entrada.getRange('D78').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_LATROC_ACU_1);
    
    Entrada.getRange('D79').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_AME_ACU_1);
    
    Entrada.getRange('D80').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_LES_CORP_ACU_1);
    
    Entrada.getRange('D81').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_HOM_ACU_1);
    
    Entrada.getRange('D82').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_MP_ACU_1);
    
    Entrada.getRange('D83').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_CRIM_SEX_ACU_1);
    
    Entrada.getRange('D84').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_ESTUP_VUL_ACU_1);
    
    Entrada.getRange('D85').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_OUT_ACU_1);
    
    //Entrada.getRange('D86').activate(); 
    //Entrada.getCurrentCell().setValue();
    
    //Entrada.getRange('D87').activate(); 
    //Entrada.getCurrentCell().setValue();
    
    //Entrada.getRange('D88').activate(); 
    //Entrada.getCurrentCell().setValue();
    
    Entrada.getRange('F43').activate(); 
    Entrada.getCurrentCell().setValue(N_ACU_2);
    
    Entrada.getRange('F44').activate(); 
    Entrada.getCurrentCell().setValue(RG_ACU_2);
    
    Entrada.getRange('F45').activate(); 
    Entrada.getCurrentCell().setValue(CPF_ACU_2);
    
    Entrada.getRange('F46').activate(); 
    Entrada.getCurrentCell().setValue(SX_ACU_2);
    
    Entrada.getRange('F47').activate(); 
    Entrada.getCurrentCell().setValue(DN_ACU_2);
    
    Entrada.getRange('F51').activate(); 
    Entrada.getCurrentCell().setValue(ORIEN_SX_ACU_2);
    
    Entrada.getRange('F52').activate(); 
    Entrada.getCurrentCell().setValue(COR_ACU_2);
    
    Entrada.getRange('F53').activate(); 
    Entrada.getCurrentCell().setValue(EC_ACU_2);
    
    Entrada.getRange('F54').activate(); 
    Entrada.getCurrentCell().setValue(UE_ACU_2);
    
    Entrada.getRange('F55').activate(); 
    Entrada.getCurrentCell().setValue(FIL_ACU_2);
    
    Entrada.getRange('F56').activate(); 
    Entrada.getCurrentCell().setValue(NAT_ACU_2);
    
    Entrada.getRange('F57').activate(); 
    Entrada.getCurrentCell().setValue(NAC_ACU_2);
    
    Entrada.getRange('F58').activate(); 
    Entrada.getCurrentCell().setValue(COND_FIS_ACU_2);
    
    Entrada.getRange('F59').activate(); 
    Entrada.getCurrentCell().setValue(ALCUNHA_ACU_2);
    
    Entrada.getRange('F60').activate(); 
    Entrada.getCurrentCell().setValue(TAT_ACU_2);
    
    Entrada.getRange('F61').activate(); 
    Entrada.getCurrentCell().setValue(ESC_ACU_2);
    
    //Entrada.getRange('F62').activate(); 
    //Entrada.getCurrentCell().setValue();
    
    Entrada.getRange('F63').activate(); 
    Entrada.getCurrentCell().setValue(END_RES_ACU_2);
    
    Entrada.getRange('F64').activate(); 
    Entrada.getCurrentCell().setValue(MUN_ACU_2);
    
    Entrada.getRange('F65').activate(); 
    Entrada.getCurrentCell().setValue(BAIRRO_ACU_2);
    
    Entrada.getRange('F66').activate(); 
    Entrada.getCurrentCell().setValue(PROF_ACU_2);
    
    Entrada.getRange('F67').activate(); 
    Entrada.getCurrentCell().setValue(END_PROF_ACU_2);
    
    Entrada.getRange('F68').activate(); 
    Entrada.getCurrentCell().setValue(MUN_PROF_ACU_2);
    
    Entrada.getRange('F69').activate(); 
    Entrada.getCurrentCell().setValue(BAIRRO_PROF_ACU_2);
    
    Entrada.getRange('F70').activate(); 
    Entrada.getCurrentCell().setValue(N_PAI_ACU_2);
    
    Entrada.getRange('F71').activate(); 
    Entrada.getCurrentCell().setValue(ESC_PAI_ACU_2);
    
    Entrada.getRange('F72').activate(); 
    Entrada.getCurrentCell().setValue(HIS_AC_PAI_ACU_2);
    
    Entrada.getRange('F73').activate(); 
    Entrada.getCurrentCell().setValue(HIS_AC_PAI_VDOM_ACU_2);
    
    //Entrada.getRange('F74').activate(); 
    //Entrada.getCurrentCell().setValue();
    
    Entrada.getRange('F75').activate(); 
    Entrada.getCurrentCell().setValue(ESC_MAE_ACU_2);
    
    Entrada.getRange('F76').activate(); 
    Entrada.getCurrentCell().setValue(N_MAE_ACU_2);
    
    Entrada.getRange('F77').activate(); 
    Entrada.getCurrentCell().setValue(HIS_AC_MAE_ACU_2);
    
    Entrada.getRange('F78').activate(); 
    Entrada.getCurrentCell().setValue(VIOL_DOM_MAE_ACU_2);
    
    //Entrada.getRange('F79').activate(); 
    //Entrada.getCurrentCell().setValue();
    
    Entrada.getRange('F80').activate(); 
    Entrada.getCurrentCell().setValue(MP_ACU_2);
    
    Entrada.getRange('F81').activate(); 
    Entrada.getCurrentCell().setValue(MI_ACU_2);
    
    Entrada.getRange('F82').activate(); 
    Entrada.getCurrentCell().setValue(HIS_MI_ACU_2);
    
    Entrada.getRange('F83').activate(); 
    Entrada.getCurrentCell().setValue(ENVTRAF_ACU_2);
    
    Entrada.getRange('F84').activate(); 
    Entrada.getCurrentCell().setValue(FOR_ACU_2);
    
    Entrada.getRange('F85').activate(); 
    Entrada.getCurrentCell().setValue(ANTEPOL_ACU_2);
    
    Entrada.getRange('F86').activate(); 
    Entrada.getCurrentCell().setValue(ANTECRIM_ACU_2);
    
    //Entrada.getRange('F87').activate(); 
    //Entrada.getCurrentCell().setValue();
    
    //Entrada.getRange('F88').activate(); 
    //Entrada.getCurrentCell().setValue();
    
    Entrada.getRange('H43').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_TRANS_ACU_2);
    
    Entrada.getRange('H44').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_DIR_HAB_ACU_2);
    
    Entrada.getRange('H45').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_FUG_RECAP_ACU_2);
    
    Entrada.getRange('H46').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ENT_POS_ACU_2);
    
    Entrada.getRange('H47').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ENT_TRAF_ACU_2);
    
    Entrada.getRange('H48').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_PORT_ARMAS_ACU_2);
    
    Entrada.getRange('H49').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_RECEP_ACU_2);
    
    Entrada.getRange('H50').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_FURTO_ACU_2);
    
    Entrada.getRange('H51').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ROUBO_PED_ACU_2);
    
    Entrada.getRange('H52').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ROUBO_BANC_ACU_2);
    
    Entrada.getRange('H53').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ROUBO_TRANSP_PUB_ACU_2);
    
    Entrada.getRange('H54').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ROUBO_TRANSP_IND_ACU_2);
    
    Entrada.getRange('H55').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ROUBO_TRANSP_IND_APP_ACU_2);
    
    Entrada.getRange('H56').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ROUBO_TRANSP_VEIC_ACU_2);
    
    Entrada.getRange('H57').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ROUBO_RESID_ACU_2);
    
    Entrada.getRange('H58').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ROUBO_OUT_ACU_2);
    
    Entrada.getRange('H59').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_LATROC_ACU_2);
    
    Entrada.getRange('H60').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_AME_ACU_2);
    
    Entrada.getRange('H61').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_LES_CORP_ACU_2);
    
    Entrada.getRange('H62').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_LES_CORP_MOR_ACU_2);
    
    Entrada.getRange('H63').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_HOM_ACU_2);
    
    Entrada.getRange('H64').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_MP_ACU_2);
    
    Entrada.getRange('H65').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_CRIM_SEX_ACU_2);
    
    Entrada.getRange('H66').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ESTUP_VUL_ACU_2);
    
    Entrada.getRange('H67').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_OUT_ACU_2);
    
    Entrada.getRange('H68').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_TRANS_ACU_2);
    
    Entrada.getRange('H69').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_FURTO_ACU_2);
    
    Entrada.getRange('H70').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_ROUBO_PED_ACU_2);
    
    Entrada.getRange('H71').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_ROUBO_BANC_ACU_2);
    
    Entrada.getRange('H72').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_ROUBO_TRANSP_PUB_ACU_2);
    
    Entrada.getRange('H73').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_ROUBO_TRANSP_IND_ACU_2);
    
    Entrada.getRange('H74').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_ROUBO_TRANSP_IND_APP_ACU_2);
    
    Entrada.getRange('H75').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_ROUBO_TRANSP_VEIC_ACU_2);
    
    Entrada.getRange('H76').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_ROUBO_RESID_ACU_2);
    
    Entrada.getRange('H77').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_ROUBO_OUT_ACU_2);
    
    Entrada.getRange('H78').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_LATROC_ACU_2);
    
    Entrada.getRange('H79').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_AME_ACU_2);
    
    Entrada.getRange('H80').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_LES_CORP_ACU_2);
    
    Entrada.getRange('H81').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_HOM_ACU_2);
    
    Entrada.getRange('H82').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_MP_ACU_2);
    
    Entrada.getRange('H83').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_CRIM_SEX_ACU_2);
    
    Entrada.getRange('H84').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_ESTUP_VUL_ACU_2);
    
    Entrada.getRange('H85').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_OUT_ACU_2);
    
    //Entrada.getRange('H86').activate(); 
    //Entrada.getCurrentCell().setValue();
    
    //Entrada.getRange('H87').activate(); 
    //Entrada.getCurrentCell().setValue();
    
    //Entrada.getRange('H88').activate(); 
    //Entrada.getCurrentCell().setValue();
    
    Entrada.getRange('B90').activate(); 
    Entrada.getCurrentCell().setValue(N_ACU_3);
    
    Entrada.getRange('B91').activate(); 
    Entrada.getCurrentCell().setValue(RG_ACU_3);
    
    Entrada.getRange('B92').activate(); 
    Entrada.getCurrentCell().setValue(CPF_ACU_3);
    
    Entrada.getRange('B93').activate(); 
    Entrada.getCurrentCell().setValue(SX_ACU_3);
    
    Entrada.getRange('B94').activate(); 
    Entrada.getCurrentCell().setValue(DN_ACU_3);
    
    Entrada.getRange('B98').activate(); 
    Entrada.getCurrentCell().setValue(ORIEN_SX_ACU_3);
    
    Entrada.getRange('B99').activate(); 
    Entrada.getCurrentCell().setValue(COR_ACU_3);
    
    Entrada.getRange('B100').activate(); 
    Entrada.getCurrentCell().setValue(EC_ACU_3);
    
    Entrada.getRange('B101').activate(); 
    Entrada.getCurrentCell().setValue(UE_ACU_3);
    
    Entrada.getRange('B102').activate(); 
    Entrada.getCurrentCell().setValue(FIL_ACU_3);
    
    Entrada.getRange('B103').activate(); 
    Entrada.getCurrentCell().setValue(NAT_ACU_3);
    
    Entrada.getRange('B104').activate(); 
    Entrada.getCurrentCell().setValue(NAC_ACU_3);
    
    Entrada.getRange('B105').activate(); 
    Entrada.getCurrentCell().setValue(COND_FIS_ACU_3);
    
    Entrada.getRange('B106').activate(); 
    Entrada.getCurrentCell().setValue(ALCUNHA_ACU_3);
    
    Entrada.getRange('B107').activate(); 
    Entrada.getCurrentCell().setValue(TAT_ACU_3);
    
    Entrada.getRange('B108').activate(); 
    Entrada.getCurrentCell().setValue(ESC_ACU_3);
    
    //Entrada.getRange('B109').activate(); 
    //Entrada.getCurrentCell().setValue();
    
    Entrada.getRange('B110').activate(); 
    Entrada.getCurrentCell().setValue(END_RES_ACU_3);
    
    Entrada.getRange('B111').activate(); 
    Entrada.getCurrentCell().setValue(MUN_ACU_3);
    
    Entrada.getRange('B112').activate(); 
    Entrada.getCurrentCell().setValue(BAIRRO_ACU_3);
    
    Entrada.getRange('B113').activate(); 
    Entrada.getCurrentCell().setValue(PROF_ACU_3);
    
    Entrada.getRange('B114').activate(); 
    Entrada.getCurrentCell().setValue(END_PROF_ACU_3);
    
    Entrada.getRange('B115').activate(); 
    Entrada.getCurrentCell().setValue(MUN_PROF_ACU_3);
    
    Entrada.getRange('B116').activate(); 
    Entrada.getCurrentCell().setValue(BAIRRO_PROF_ACU_3);
    
    Entrada.getRange('B117').activate(); 
    Entrada.getCurrentCell().setValue(N_PAI_ACU_3);
    
    Entrada.getRange('B118').activate(); 
    Entrada.getCurrentCell().setValue(ESC_PAI_ACU_3);
    
    Entrada.getRange('B119').activate(); 
    Entrada.getCurrentCell().setValue(HIS_AC_PAI_ACU_3);
    
    Entrada.getRange('B120').activate(); 
    Entrada.getCurrentCell().setValue(HIS_AC_PAI_VDOM_ACU_3);
    
    //Entrada.getRange('B121').activate(); 
    //Entrada.getCurrentCell().setValue();
    
    Entrada.getRange('B122').activate(); 
    Entrada.getCurrentCell().setValue(ESC_MAE_ACU_3);
    
    Entrada.getRange('B123').activate(); 
    Entrada.getCurrentCell().setValue(N_MAE_ACU_3);
    
    Entrada.getRange('B124').activate(); 
    Entrada.getCurrentCell().setValue(HIS_AC_MAE_ACU_3);
    
    Entrada.getRange('B125').activate(); 
    Entrada.getCurrentCell().setValue(VIOL_DOM_MAE_ACU_3);
    
    //Entrada.getRange('B126').activate(); 
    //Entrada.getCurrentCell().setValue();
    
    Entrada.getRange('B127').activate(); 
    Entrada.getCurrentCell().setValue(MP_ACU_3);
    
    Entrada.getRange('B128').activate(); 
    Entrada.getCurrentCell().setValue(MI_ACU_3);
    
    Entrada.getRange('B129').activate(); 
    Entrada.getCurrentCell().setValue(HIS_MI_ACU_3);
    
    Entrada.getRange('B130').activate(); 
    Entrada.getCurrentCell().setValue(ENVTRAF_ACU_3);
    
    Entrada.getRange('B131').activate(); 
    Entrada.getCurrentCell().setValue(FOR_ACU_3);
    
    Entrada.getRange('B132').activate(); 
    Entrada.getCurrentCell().setValue(ANTEPOL_ACU_3);
    
    Entrada.getRange('B133').activate(); 
    Entrada.getCurrentCell().setValue(ANTECRIM_ACU_3);
    
    //Entrada.getRange('B134').activate(); 
    //Entrada.getCurrentCell().setValue();
    
    //Entrada.getRange('B135').activate(); 
    //Entrada.getCurrentCell().setValue();
    
    Entrada.getRange('D90').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_TRANS_ACU_3);
    
    Entrada.getRange('D91').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_DIR_HAB_ACU_3);
    
    Entrada.getRange('D92').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_FUG_RECAP_ACU_3);
    
    Entrada.getRange('D93').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ENT_POS_ACU_3);
    
    Entrada.getRange('D94').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ENT_TRAF_ACU_3);
    
    Entrada.getRange('D95').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_PORT_ARMAS_ACU_3);
    
    Entrada.getRange('D96').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_RECEP_ACU_3);
    
    Entrada.getRange('D97').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_FURTO_ACU_3);
    
    Entrada.getRange('D98').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ROUBO_PED_ACU_3);
    
    Entrada.getRange('D99').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ROUBO_BANC_ACU_3);
    
    Entrada.getRange('D100').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ROUBO_TRANSP_PUB_ACU_3);
    
    Entrada.getRange('D101').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ROUBO_TRANSP_IND_ACU_3);
    
    Entrada.getRange('D102').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ROUBO_TRANSP_IND_APP_ACU_3);
    
    Entrada.getRange('D103').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ROUBO_TRANSP_VEIC_ACU_3);
    
    Entrada.getRange('D104').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ROUBO_RESID_ACU_3);
    
    Entrada.getRange('D105').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ROUBO_OUT_ACU_3);
    
    Entrada.getRange('D106').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_LATROC_ACU_3);
    
    Entrada.getRange('D107').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_AME_ACU_3);
    
    Entrada.getRange('D108').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_LES_CORP_ACU_3);
    
    Entrada.getRange('D109').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_LES_CORP_MOR_ACU_3);
    
    Entrada.getRange('D110').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_HOM_ACU_3);
    
    Entrada.getRange('D111').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_MP_ACU_3);
    
    Entrada.getRange('D112').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_CRIM_SEX_ACU_3);
    
    Entrada.getRange('D113').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ESTUP_VUL_ACU_3);
    
    Entrada.getRange('D114').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_OUT_ACU_3);
    
    Entrada.getRange('D115').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_TRANS_ACU_3);
    
    Entrada.getRange('D116').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_FURTO_ACU_3);
    
    Entrada.getRange('D117').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_ROUBO_PED_ACU_3);
    
    Entrada.getRange('D118').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_ROUBO_BANC_ACU_3);
    
    Entrada.getRange('D119').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_ROUBO_TRANSP_PUB_ACU_3);
    
    Entrada.getRange('D120').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_ROUBO_TRANSP_IND_ACU_3);
    
    Entrada.getRange('D121').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_ROUBO_TRANSP_IND_APP_ACU_3);
    
    Entrada.getRange('D122').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_ROUBO_TRANSP_VEIC_ACU_3);
    
    Entrada.getRange('D123').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_ROUBO_RESID_ACU_3);
    
    Entrada.getRange('D124').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_ROUBO_OUT_ACU_3);
    
    Entrada.getRange('D125').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_LATROC_ACU_3);
    
    Entrada.getRange('D126').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_AME_ACU_3);
    
    Entrada.getRange('D127').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_LES_CORP_ACU_3);
    
    Entrada.getRange('D128').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_HOM_ACU_3);
    
    Entrada.getRange('D129').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_MP_ACU_3);
    
    Entrada.getRange('D130').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_CRIM_SEX_ACU_3);
    
    Entrada.getRange('D131').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_ESTUP_VUL_ACU_3);
    
    Entrada.getRange('D132').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_OUT_ACU_3);
    
    //Entrada.getRange('D133').activate(); 
    //Entrada.getCurrentCell().setValue();
    
    //Entrada.getRange('D134').activate(); 
    //Entrada.getCurrentCell().setValue();
    
    //Entrada.getRange('D135').activate(); 
    //Entrada.getCurrentCell().setValue();
    
    Entrada.getRange('F90').activate(); 
    Entrada.getCurrentCell().setValue(N_ACU_4);
    
    Entrada.getRange('F91').activate(); 
    Entrada.getCurrentCell().setValue(RG_ACU_4);
    
    Entrada.getRange('F92').activate(); 
    Entrada.getCurrentCell().setValue(CPF_ACU_4);
    
    Entrada.getRange('F93').activate(); 
    Entrada.getCurrentCell().setValue(SX_ACU_4);
    
    Entrada.getRange('F94').activate(); 
    Entrada.getCurrentCell().setValue(DN_ACU_4);
    
    Entrada.getRange('F98').activate(); 
    Entrada.getCurrentCell().setValue(ORIEN_SX_ACU_4);
    
    Entrada.getRange('F99').activate(); 
    Entrada.getCurrentCell().setValue(COR_ACU_4);
    
    Entrada.getRange('F100').activate(); 
    Entrada.getCurrentCell().setValue(EC_ACU_4);
    
    Entrada.getRange('F101').activate(); 
    Entrada.getCurrentCell().setValue(UE_ACU_4);
    
    Entrada.getRange('F102').activate(); 
    Entrada.getCurrentCell().setValue(FIL_ACU_4);
    
    Entrada.getRange('F103').activate(); 
    Entrada.getCurrentCell().setValue(NAT_ACU_4);
    
    Entrada.getRange('F104').activate(); 
    Entrada.getCurrentCell().setValue(NAC_ACU_4);
    
    Entrada.getRange('F105').activate(); 
    Entrada.getCurrentCell().setValue(COND_FIS_ACU_4);
    
    Entrada.getRange('F106').activate(); 
    Entrada.getCurrentCell().setValue(ALCUNHA_ACU_4);
    
    Entrada.getRange('F107').activate(); 
    Entrada.getCurrentCell().setValue(TAT_ACU_4);
    
    Entrada.getRange('F108').activate(); 
    Entrada.getCurrentCell().setValue(ESC_ACU_4);
    
    //Entrada.getRange('F109').activate(); 
    //Entrada.getCurrentCell().setValue();
    
    Entrada.getRange('F110').activate(); 
    Entrada.getCurrentCell().setValue(END_RES_ACU_4);
    
    Entrada.getRange('F111').activate(); 
    Entrada.getCurrentCell().setValue(MUN_ACU_4);
    
    Entrada.getRange('F112').activate(); 
    Entrada.getCurrentCell().setValue(BAIRRO_ACU_4);
    
    Entrada.getRange('F113').activate(); 
    Entrada.getCurrentCell().setValue(PROF_ACU_4);
    
    Entrada.getRange('F114').activate(); 
    Entrada.getCurrentCell().setValue(END_PROF_ACU_4);
    
    Entrada.getRange('F115').activate(); 
    Entrada.getCurrentCell().setValue(MUN_PROF_ACU_4);
    
    Entrada.getRange('F116').activate(); 
    Entrada.getCurrentCell().setValue(BAIRRO_PROF_ACU_4);
    
    Entrada.getRange('F117').activate(); 
    Entrada.getCurrentCell().setValue(N_PAI_ACU_4);
    
    Entrada.getRange('F118').activate(); 
    Entrada.getCurrentCell().setValue(ESC_PAI_ACU_4);
    
    Entrada.getRange('F119').activate(); 
    Entrada.getCurrentCell().setValue(HIS_AC_PAI_ACU_4);
    
    Entrada.getRange('F120').activate(); 
    Entrada.getCurrentCell().setValue(HIS_AC_PAI_VDOM_ACU_4);
    
    //Entrada.getRange('F121').activate(); 
    //Entrada.getCurrentCell().setValue();
    
    Entrada.getRange('F122').activate(); 
    Entrada.getCurrentCell().setValue(ESC_MAE_ACU_4);
    
    Entrada.getRange('F123').activate(); 
    Entrada.getCurrentCell().setValue(N_MAE_ACU_4);
    
    Entrada.getRange('F124').activate(); 
    Entrada.getCurrentCell().setValue(HIS_AC_MAE_ACU_4);
    
    Entrada.getRange('F125').activate(); 
    Entrada.getCurrentCell().setValue(VIOL_DOM_MAE_ACU_4);
    
    //Entrada.getRange('F126').activate(); 
    //Entrada.getCurrentCell().setValue();
    
    Entrada.getRange('F127').activate(); 
    Entrada.getCurrentCell().setValue(MP_ACU_4);
    
    Entrada.getRange('F128').activate(); 
    Entrada.getCurrentCell().setValue(MI_ACU_4);
    
    Entrada.getRange('F129').activate(); 
    Entrada.getCurrentCell().setValue(HIS_MI_ACU_4);
    
    Entrada.getRange('F130').activate(); 
    Entrada.getCurrentCell().setValue(ENVTRAF_ACU_4);
    
    Entrada.getRange('F131').activate(); 
    Entrada.getCurrentCell().setValue(FOR_ACU_4);
    
    Entrada.getRange('F132').activate(); 
    Entrada.getCurrentCell().setValue(ANTEPOL_ACU_4);
    
    Entrada.getRange('F133').activate(); 
    Entrada.getCurrentCell().setValue(ANTECRIM_ACU_4);
    
    //Entrada.getRange('F134').activate(); 
    //Entrada.getCurrentCell().setValue();
    
    //Entrada.getRange('F135').activate(); 
    //Entrada.getCurrentCell().setValue();
    
    Entrada.getRange('H90').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_TRANS_ACU_4);
    
    Entrada.getRange('H91').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_DIR_HAB_ACU_4);
    
    Entrada.getRange('H92').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_FUG_RECAP_ACU_4);
    
    Entrada.getRange('H93').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ENT_POS_ACU_4);
    
    Entrada.getRange('H94').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ENT_TRAF_ACU_4);
    
    Entrada.getRange('H95').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_PORT_ARMAS_ACU_4);
    
    Entrada.getRange('H96').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_RECEP_ACU_4);
    
    Entrada.getRange('H97').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_FURTO_ACU_4);
    
    Entrada.getRange('H98').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ROUBO_PED_ACU_4);
    
    Entrada.getRange('H99').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ROUBO_BANC_ACU_4);
    
    Entrada.getRange('H100').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ROUBO_TRANSP_PUB_ACU_4);
    
    Entrada.getRange('H101').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ROUBO_TRANSP_IND_ACU_4);
    
    Entrada.getRange('H102').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ROUBO_TRANSP_IND_APP_ACU_4);
    
    Entrada.getRange('H103').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ROUBO_TRANSP_VEIC_ACU_4);
    
    Entrada.getRange('H104').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ROUBO_RESID_ACU_4);
    
    Entrada.getRange('H105').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ROUBO_OUT_ACU_4);
    
    Entrada.getRange('H106').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_LATROC_ACU_4);
    
    Entrada.getRange('H107').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_AME_ACU_4);
    
    Entrada.getRange('H108').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_LES_CORP_ACU_4);
    
    Entrada.getRange('H109').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_LES_CORP_MOR_ACU_4);
    
    Entrada.getRange('H110').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_HOM_ACU_4);
    
    Entrada.getRange('H111').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_MP_ACU_4);
    
    Entrada.getRange('H112').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_CRIM_SEX_ACU_4);
    
    Entrada.getRange('H113').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ESTUP_VUL_ACU_4);
    
    Entrada.getRange('H114').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_OUT_ACU_4);
    
    Entrada.getRange('H115').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_TRANS_ACU_4);
    
    Entrada.getRange('H116').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_FURTO_ACU_4);
    
    Entrada.getRange('H117').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_ROUBO_PED_ACU_4);
    
    Entrada.getRange('H118').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_ROUBO_BANC_ACU_4);
    
    Entrada.getRange('H119').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_ROUBO_TRANSP_PUB_ACU_4);
    
    Entrada.getRange('H120').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_ROUBO_TRANSP_IND_ACU_4);
    
    Entrada.getRange('H121').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_ROUBO_TRANSP_IND_APP_ACU_4);
    
    Entrada.getRange('H122').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_ROUBO_TRANSP_VEIC_ACU_4);
    
    Entrada.getRange('H123').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_ROUBO_RESID_ACU_4);
    
    Entrada.getRange('H124').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_ROUBO_OUT_ACU_4);
    
    Entrada.getRange('H125').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_LATROC_ACU_4);
    
    Entrada.getRange('H126').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_AME_ACU_4);
    
    Entrada.getRange('H127').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_LES_CORP_ACU_4);
    
    Entrada.getRange('H128').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_HOM_ACU_4);
    
    Entrada.getRange('H129').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_MP_ACU_4);
    
    Entrada.getRange('H130').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_CRIM_SEX_ACU_4);
    
    Entrada.getRange('H131').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_ESTUP_VUL_ACU_4);
    
    Entrada.getRange('H132').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_OUT_ACU_4);
    
    //Entrada.getRange('H133').activate(); 
    //Entrada.getCurrentCell().setValue();
    
    //Entrada.getRange('H134').activate(); 
    //Entrada.getCurrentCell().setValue();
    
    //Entrada.getRange('H135').activate(); 
    //Entrada.getCurrentCell().setValue();
    
    Entrada.getRange('B137').activate(); 
    Entrada.getCurrentCell().setValue(N_ACU_5);
    
    Entrada.getRange('B138').activate(); 
    Entrada.getCurrentCell().setValue(RG_ACU_5);
    
    Entrada.getRange('B139').activate(); 
    Entrada.getCurrentCell().setValue(CPF_ACU_5);
    
    Entrada.getRange('B140').activate(); 
    Entrada.getCurrentCell().setValue(SX_ACU_5);
    
    Entrada.getRange('B141').activate(); 
    Entrada.getCurrentCell().setValue(DN_ACU_5);
    
    Entrada.getRange('B145').activate(); 
    Entrada.getCurrentCell().setValue(ORIEN_SX_ACU_5);
    
    Entrada.getRange('B146').activate(); 
    Entrada.getCurrentCell().setValue(COR_ACU_5);
    
    Entrada.getRange('B147').activate(); 
    Entrada.getCurrentCell().setValue(EC_ACU_5);
    
    Entrada.getRange('B148').activate(); 
    Entrada.getCurrentCell().setValue(UE_ACU_5);
    
    Entrada.getRange('B149').activate(); 
    Entrada.getCurrentCell().setValue(FIL_ACU_5);
    
    Entrada.getRange('B150').activate(); 
    Entrada.getCurrentCell().setValue(NAT_ACU_5);
    
    Entrada.getRange('B151').activate(); 
    Entrada.getCurrentCell().setValue(NAC_ACU_5);
    
    Entrada.getRange('B152').activate(); 
    Entrada.getCurrentCell().setValue(COND_FIS_ACU_5);
    
    Entrada.getRange('B153').activate(); 
    Entrada.getCurrentCell().setValue(ALCUNHA_ACU_5);
    
    Entrada.getRange('B154').activate(); 
    Entrada.getCurrentCell().setValue(TAT_ACU_5);
    
    Entrada.getRange('B155').activate(); 
    Entrada.getCurrentCell().setValue(ESC_ACU_5);
    
    //Entrada.getRange('B156').activate(); 
    //Entrada.getCurrentCell().setValue();
    
    Entrada.getRange('B157').activate(); 
    Entrada.getCurrentCell().setValue(END_RES_ACU_5);
    
    Entrada.getRange('B158').activate(); 
    Entrada.getCurrentCell().setValue(MUN_ACU_5);
    
    Entrada.getRange('B159').activate(); 
    Entrada.getCurrentCell().setValue(BAIRRO_ACU_5);
    
    Entrada.getRange('B160').activate(); 
    Entrada.getCurrentCell().setValue(PROF_ACU_5);
    
    Entrada.getRange('B161').activate(); 
    Entrada.getCurrentCell().setValue(END_PROF_ACU_5);
    
    Entrada.getRange('B162').activate(); 
    Entrada.getCurrentCell().setValue(MUN_PROF_ACU_5);
    
    Entrada.getRange('B163').activate(); 
    Entrada.getCurrentCell().setValue(BAIRRO_PROF_ACU_5);
    
    Entrada.getRange('B164').activate(); 
    Entrada.getCurrentCell().setValue(N_PAI_ACU_5);
    
    Entrada.getRange('B165').activate(); 
    Entrada.getCurrentCell().setValue(ESC_PAI_ACU_5);
    
    Entrada.getRange('B166').activate(); 
    Entrada.getCurrentCell().setValue(HIS_AC_PAI_ACU_5);
    
    Entrada.getRange('B167').activate(); 
    Entrada.getCurrentCell().setValue(HIS_AC_PAI_VDOM_ACU_5);
    
    //Entrada.getRange('B168').activate(); 
    //Entrada.getCurrentCell().setValue();
    
    Entrada.getRange('B169').activate(); 
    Entrada.getCurrentCell().setValue(ESC_MAE_ACU_5);
    
    Entrada.getRange('B170').activate(); 
    Entrada.getCurrentCell().setValue(N_MAE_ACU_5);
    
    Entrada.getRange('B171').activate(); 
    Entrada.getCurrentCell().setValue(HIS_AC_MAE_ACU_5);
    
    Entrada.getRange('B172').activate(); 
    Entrada.getCurrentCell().setValue(VIOL_DOM_MAE_ACU_5);
    
    //Entrada.getRange('B173').activate(); 
    //Entrada.getCurrentCell().setValue();
    
    Entrada.getRange('B174').activate(); 
    Entrada.getCurrentCell().setValue(MP_ACU_5);
    
    Entrada.getRange('B175').activate(); 
    Entrada.getCurrentCell().setValue(MI_ACU_5);
    
    Entrada.getRange('B176').activate(); 
    Entrada.getCurrentCell().setValue(HIS_MI_ACU_5);
    
    Entrada.getRange('B177').activate(); 
    Entrada.getCurrentCell().setValue(ENVTRAF_ACU_5);
    
    Entrada.getRange('B178').activate(); 
    Entrada.getCurrentCell().setValue(FOR_ACU_5);
    
    Entrada.getRange('B179').activate(); 
    Entrada.getCurrentCell().setValue(ANTEPOL_ACU_5);
    
    Entrada.getRange('B180').activate(); 
    Entrada.getCurrentCell().setValue(ANTECRIM_ACU_5);
    
    //Entrada.getRange('B181').activate(); 
    //Entrada.getCurrentCell().setValue();
    
    //Entrada.getRange('B182').activate(); 
    //Entrada.getCurrentCell().setValue();
    
    Entrada.getRange('D137').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_TRANS_ACU_5);
    
    Entrada.getRange('D138').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_DIR_HAB_ACU_5);
    
    Entrada.getRange('D139').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_FUG_RECAP_ACU_5);
    
    Entrada.getRange('D140').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ENT_POS_ACU_5);
    
    Entrada.getRange('D141').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ENT_TRAF_ACU_5);
    
    Entrada.getRange('D142').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_PORT_ARMAS_ACU_5);
    
    Entrada.getRange('D143').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_RECEP_ACU_5);
    
    Entrada.getRange('D144').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_FURTO_ACU_5);
    
    Entrada.getRange('D145').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ROUBO_PED_ACU_5);
    
    Entrada.getRange('D146').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ROUBO_BANC_ACU_5);
    
    Entrada.getRange('D147').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ROUBO_TRANSP_PUB_ACU_5);
    
    Entrada.getRange('D148').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ROUBO_TRANSP_IND_ACU_5);
    
    Entrada.getRange('D149').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ROUBO_TRANSP_IND_APP_ACU_5);
    
    Entrada.getRange('D150').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ROUBO_TRANSP_VEIC_ACU_5);
    
    Entrada.getRange('D151').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ROUBO_RESID_ACU_5);
    
    Entrada.getRange('D152').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ROUBO_OUT_ACU_5);
    
    Entrada.getRange('D153').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_LATROC_ACU_5);
    
    Entrada.getRange('D154').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_AME_ACU_5);
    
    Entrada.getRange('D155').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_LES_CORP_ACU_5);
    
    Entrada.getRange('D156').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_LES_CORP_MOR_ACU_5);
    
    Entrada.getRange('D157').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_HOM_ACU_5);
    
    Entrada.getRange('D158').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_MP_ACU_5);
    
    Entrada.getRange('D159').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_CRIM_SEX_ACU_5);
    
    Entrada.getRange('D160').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ESTUP_VUL_ACU_5);
    
    Entrada.getRange('D161').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_OUT_ACU_5);
    
    Entrada.getRange('D162').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_TRANS_ACU_5);
    
    Entrada.getRange('D163').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_FURTO_ACU_5);
    
    Entrada.getRange('D164').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_ROUBO_PED_ACU_5);
    
    Entrada.getRange('D165').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_ROUBO_BANC_ACU_5);
    
    Entrada.getRange('D166').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_ROUBO_TRANSP_PUB_ACU_5);
    
    Entrada.getRange('D167').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_ROUBO_TRANSP_IND_ACU_5);
    
    Entrada.getRange('D168').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_ROUBO_TRANSP_IND_APP_ACU_5);
    
    Entrada.getRange('D169').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_ROUBO_TRANSP_VEIC_ACU_5);
    
    Entrada.getRange('D170').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_ROUBO_RESID_ACU_5);
    
    Entrada.getRange('D171').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_ROUBO_OUT_ACU_5);
    
    Entrada.getRange('D172').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_LATROC_ACU_5);
    
    Entrada.getRange('D173').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_AME_ACU_5);
    
    Entrada.getRange('D174').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_LES_CORP_ACU_5);
    
    Entrada.getRange('D175').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_HOM_ACU_5);
    
    Entrada.getRange('D176').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_MP_ACU_5);
    
    Entrada.getRange('D177').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_CRIM_SEX_ACU_5);
    
    Entrada.getRange('D178').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_ESTUP_VUL_ACU_5);
    
    Entrada.getRange('D179').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_OUT_ACU_5);
    
    //Entrada.getRange('D180').activate(); 
    //Entrada.getCurrentCell().setValue();
    
    //Entrada.getRange('D181').activate(); 
    //Entrada.getCurrentCell().setValue();
    
    //Entrada.getRange('D182').activate(); 
    //Entrada.getCurrentCell().setValue();
    
    Entrada.getRange('F137').activate(); 
    Entrada.getCurrentCell().setValue(N_ACU_6);
    
    Entrada.getRange('F138').activate(); 
    Entrada.getCurrentCell().setValue(RG_ACU_6);
    
    Entrada.getRange('F139').activate(); 
    Entrada.getCurrentCell().setValue(CPF_ACU_6);
    
    Entrada.getRange('F140').activate(); 
    Entrada.getCurrentCell().setValue(SX_ACU_6);
    
    Entrada.getRange('F141').activate(); 
    Entrada.getCurrentCell().setValue(DN_ACU_6);
    
    Entrada.getRange('F145').activate(); 
    Entrada.getCurrentCell().setValue(ORIEN_SX_ACU_6);
    
    Entrada.getRange('F146').activate(); 
    Entrada.getCurrentCell().setValue(COR_ACU_6);
    
    Entrada.getRange('F147').activate(); 
    Entrada.getCurrentCell().setValue(EC_ACU_6);
    
    Entrada.getRange('F148').activate(); 
    Entrada.getCurrentCell().setValue(UE_ACU_6);
    
    Entrada.getRange('F149').activate(); 
    Entrada.getCurrentCell().setValue(FIL_ACU_6);
    
    Entrada.getRange('F150').activate(); 
    Entrada.getCurrentCell().setValue(NAT_ACU_6);
    
    Entrada.getRange('F151').activate(); 
    Entrada.getCurrentCell().setValue(NAC_ACU_6);
    
    Entrada.getRange('F152').activate(); 
    Entrada.getCurrentCell().setValue(COND_FIS_ACU_6);
    
    Entrada.getRange('F153').activate(); 
    Entrada.getCurrentCell().setValue(ALCUNHA_ACU_6);
    
    Entrada.getRange('F154').activate(); 
    Entrada.getCurrentCell().setValue(TAT_ACU_6);
    
    Entrada.getRange('F155').activate(); 
    Entrada.getCurrentCell().setValue(ESC_ACU_6);
    
    //Entrada.getRange('F156').activate(); 
    //Entrada.getCurrentCell().setValue();
    
    Entrada.getRange('F157').activate(); 
    Entrada.getCurrentCell().setValue(END_RES_ACU_6);
    
    Entrada.getRange('F158').activate(); 
    Entrada.getCurrentCell().setValue(MUN_ACU_6);
    
    Entrada.getRange('F159').activate(); 
    Entrada.getCurrentCell().setValue(BAIRRO_ACU_6);
    
    Entrada.getRange('F160').activate(); 
    Entrada.getCurrentCell().setValue(PROF_ACU_6);
    
    Entrada.getRange('F161').activate(); 
    Entrada.getCurrentCell().setValue(END_PROF_ACU_6);
    
    Entrada.getRange('F162').activate(); 
    Entrada.getCurrentCell().setValue(MUN_PROF_ACU_6);
    
    Entrada.getRange('F163').activate(); 
    Entrada.getCurrentCell().setValue(BAIRRO_PROF_ACU_6);
    
    Entrada.getRange('F164').activate(); 
    Entrada.getCurrentCell().setValue(N_PAI_ACU_6);
    
    Entrada.getRange('F165').activate(); 
    Entrada.getCurrentCell().setValue(ESC_PAI_ACU_6);
    
    Entrada.getRange('F166').activate(); 
    Entrada.getCurrentCell().setValue(HIS_AC_PAI_ACU_6);
    
    Entrada.getRange('F167').activate(); 
    Entrada.getCurrentCell().setValue(HIS_AC_PAI_VDOM_ACU_6);
    
    //Entrada.getRange('F168').activate(); 
    //Entrada.getCurrentCell().setValue();
    
    Entrada.getRange('F169').activate(); 
    Entrada.getCurrentCell().setValue(ESC_MAE_ACU_6);
    
    Entrada.getRange('F170').activate(); 
    Entrada.getCurrentCell().setValue(N_MAE_ACU_6);
    
    Entrada.getRange('F171').activate(); 
    Entrada.getCurrentCell().setValue(HIS_AC_MAE_ACU_6);
    
    Entrada.getRange('F172').activate(); 
    Entrada.getCurrentCell().setValue(VIOL_DOM_MAE_ACU_6);
    
    //Entrada.getRange('F173').activate(); 
    //Entrada.getCurrentCell().setValue();
    
    Entrada.getRange('F174').activate(); 
    Entrada.getCurrentCell().setValue(MP_ACU_6);
    
    Entrada.getRange('F175').activate(); 
    Entrada.getCurrentCell().setValue(MI_ACU_6);
    
    Entrada.getRange('F176').activate(); 
    Entrada.getCurrentCell().setValue(HIS_MI_ACU_6);
    
    Entrada.getRange('F177').activate(); 
    Entrada.getCurrentCell().setValue(ENVTRAF_ACU_6);
    
    Entrada.getRange('F178').activate(); 
    Entrada.getCurrentCell().setValue(FOR_ACU_6);
    
    Entrada.getRange('F179').activate(); 
    Entrada.getCurrentCell().setValue(ANTEPOL_ACU_6);
    
    Entrada.getRange('F180').activate(); 
    Entrada.getCurrentCell().setValue(ANTECRIM_ACU_6);
    
    //Entrada.getRange('F181').activate(); 
    //Entrada.getCurrentCell().setValue();
    
    //Entrada.getRange('F182').activate(); 
    //Entrada.getCurrentCell().setValue();
    
    Entrada.getRange('H137').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_TRANS_ACU_6);
    
    Entrada.getRange('H138').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_DIR_HAB_ACU_6);
    
    Entrada.getRange('H139').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_FUG_RECAP_ACU_6);
    
    Entrada.getRange('H140').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ENT_POS_ACU_6);
    
    Entrada.getRange('H141').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ENT_TRAF_ACU_6);
    
    Entrada.getRange('H142').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_PORT_ARMAS_ACU_6);
    
    Entrada.getRange('H143').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_RECEP_ACU_6);
    
    Entrada.getRange('H144').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_FURTO_ACU_6);
    
    Entrada.getRange('H145').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ROUBO_PED_ACU_6);
    
    Entrada.getRange('H146').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ROUBO_BANC_ACU_6);
    
    Entrada.getRange('H147').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ROUBO_TRANSP_PUB_ACU_6);
    
    Entrada.getRange('H148').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ROUBO_TRANSP_IND_ACU_6);
    
    Entrada.getRange('H149').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ROUBO_TRANSP_IND_APP_ACU_6);
    
    Entrada.getRange('H150').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ROUBO_TRANSP_VEIC_ACU_6);
    
    Entrada.getRange('H151').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ROUBO_RESID_ACU_6);
    
    Entrada.getRange('H152').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ROUBO_OUT_ACU_6);
    
    Entrada.getRange('H153').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_LATROC_ACU_6);
    
    Entrada.getRange('H154').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_AME_ACU_6);
    
    Entrada.getRange('H155').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_LES_CORP_ACU_6);
    
    Entrada.getRange('H156').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_LES_CORP_MOR_ACU_6);
    
    Entrada.getRange('H157').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_HOM_ACU_6);
    
    Entrada.getRange('H158').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_MP_ACU_6);
    
    Entrada.getRange('H159').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_CRIM_SEX_ACU_6);
    
    Entrada.getRange('H160').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_ESTUP_VUL_ACU_6);
    
    Entrada.getRange('H161').activate(); 
    Entrada.getCurrentCell().setValue(ANTEC_OUT_ACU_6);
    
    Entrada.getRange('H162').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_TRANS_ACU_6);
    
    Entrada.getRange('H163').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_FURTO_ACU_6);
    
    Entrada.getRange('H164').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_ROUBO_PED_ACU_6);
    
    Entrada.getRange('H165').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_ROUBO_BANC_ACU_6);
    
    Entrada.getRange('H166').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_ROUBO_TRANSP_PUB_ACU_6);
    
    Entrada.getRange('H167').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_ROUBO_TRANSP_IND_ACU_6);
    
    Entrada.getRange('H168').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_ROUBO_TRANSP_IND_APP_ACU_6);
    
    Entrada.getRange('H169').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_ROUBO_TRANSP_VEIC_ACU_6);
    
    Entrada.getRange('H170').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_ROUBO_RESID_ACU_6);
    
    Entrada.getRange('H171').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_ROUBO_OUT_ACU_6);
    
    Entrada.getRange('H172').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_LATROC_ACU_6);
    
    Entrada.getRange('H173').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_AME_ACU_6);
    
    Entrada.getRange('H174').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_LES_CORP_ACU_6);
    
    Entrada.getRange('H175').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_HOM_ACU_6);
    
    Entrada.getRange('H176').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_MP_ACU_6);
    
    Entrada.getRange('H177').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_CRIM_SEX_ACU_6);
    
    Entrada.getRange('H178').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_ESTUP_VUL_ACU_6);
    
    Entrada.getRange('H179').activate(); 
    Entrada.getCurrentCell().setValue(HIST_VIT_OUT_ACU_6);
    
    //Entrada.getRange('H180').activate(); 
    //Entrada.getCurrentCell().setValue();
    
    //Entrada.getRange('H181').activate(); 
    //Entrada.getCurrentCell().setValue();
    
    //Entrada.getRange('H182').activate(); 
    //Entrada.getCurrentCell().setValue();

  }else{
   Browser.msgBox("Ocorrência não localizada!") 
  
  }

}

function EditarOcorrencia(){ 
  
  var Entrada = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Entrada');
  var Pesquisa = Entrada.getRange('B2').getValue();
  var SheetBanco = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Banco');
  SheetBanco.getActiveCell();
  
  var LocalPesquisa = SheetBanco.getRange(2, 2, SheetBanco.getLastRow()).getValues(); 
  var Resultado = LocalPesquisa.Pesquisa(Pesquisa);
  
  var Linha = Resultado + 2;
  var verif = Verificar();
  
  if(Resultado != -1){
   if(verif == true){
     SheetBanco.getActiveCell();
    
      SheetBanco.getRange(Linha, 2).setValue(Entrada.getRange('B2').getValue());
      SheetBanco.getRange(Linha, 3).setValue(Entrada.getRange('B3').getValue());
      SheetBanco.getRange(Linha, 4).setValue(Entrada.getRange('B4').getValue());
      SheetBanco.getRange(Linha, 5).setValue(Entrada.getRange('B5').getValue());
      SheetBanco.getRange(Linha, 6).setValue(Entrada.getRange('B6').getValue());
      SheetBanco.getRange(Linha, 7).setValue(Entrada.getRange('B7').getValue());
      SheetBanco.getRange(Linha, 8).setValue(Entrada.getRange('B8').getValue());
      SheetBanco.getRange(Linha, 9).setValue(Entrada.getRange('B9').getValue());
      SheetBanco.getRange(Linha, 10).setValue(Entrada.getRange('B10').getValue());
      SheetBanco.getRange(Linha, 11).setValue(Entrada.getRange('B11').getValue());
      SheetBanco.getRange(Linha, 12).setValue(Entrada.getRange('B12').getValue());
      SheetBanco.getRange(Linha, 13).setValue(Entrada.getRange('B13').getValue());
      SheetBanco.getRange(Linha, 14).setValue(Entrada.getRange('B14').getValue());
      SheetBanco.getRange(Linha, 15).setValue(Entrada.getRange('B15').getValue());
      SheetBanco.getRange(Linha, 16).setValue(Entrada.getRange('B16').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('B17').getValue());
      SheetBanco.getRange(Linha, 17).setValue(Entrada.getRange('D2').getValue());
      SheetBanco.getRange(Linha, 18).setValue(Entrada.getRange('D3').getValue());
      SheetBanco.getRange(Linha, 19).setValue(Entrada.getRange('D4').getValue());
      SheetBanco.getRange(Linha, 20).setValue(Entrada.getRange('D5').getValue());
      SheetBanco.getRange(Linha, 21).setValue(Entrada.getRange('D6').getValue());
      SheetBanco.getRange(Linha, 22).setValue(Entrada.getRange('D7').getValue());
      SheetBanco.getRange(Linha, 23).setValue(Entrada.getRange('D8').getValue());
      SheetBanco.getRange(Linha, 24).setValue(Entrada.getRange('D9').getValue());
      SheetBanco.getRange(Linha, 25).setValue(Entrada.getRange('D10').getValue());
      SheetBanco.getRange(Linha, 26).setValue(Entrada.getRange('D11').getValue());
      SheetBanco.getRange(Linha, 27).setValue(Entrada.getRange('D12').getValue());
      SheetBanco.getRange(Linha, 28).setValue(Entrada.getRange('D13').getValue());
      SheetBanco.getRange(Linha, 29).setValue(Entrada.getRange('D14').getValue());
      SheetBanco.getRange(Linha, 30).setValue(Entrada.getRange('D15').getValue());
      SheetBanco.getRange(Linha, 31).setValue(Entrada.getRange('D16').getValue());
      SheetBanco.getRange(Linha, 32).setValue(Entrada.getRange('D17').getValue());
      SheetBanco.getRange(Linha, 33).setValue(Entrada.getRange('F2').getValue());
      SheetBanco.getRange(Linha, 34).setValue(Entrada.getRange('F3').getValue());
      SheetBanco.getRange(Linha, 35).setValue(Entrada.getRange('F4').getValue());
      SheetBanco.getRange(Linha, 36).setValue(Entrada.getRange('F5').getValue());
      SheetBanco.getRange(Linha, 37).setValue(Entrada.getRange('F6').getValue());
      SheetBanco.getRange(Linha, 38).setValue(Entrada.getRange('F7').getValue());
      SheetBanco.getRange(Linha, 39).setValue(Entrada.getRange('F8').getValue());
      SheetBanco.getRange(Linha, 40).setValue(Entrada.getRange('F9').getValue());
      SheetBanco.getRange(Linha, 41).setValue(Entrada.getRange('F10').getValue());
      SheetBanco.getRange(Linha, 42).setValue(Entrada.getRange('F11').getValue());
      SheetBanco.getRange(Linha, 43).setValue(Entrada.getRange('F12').getValue());
      SheetBanco.getRange(Linha, 44).setValue(Entrada.getRange('F13').getValue());
      SheetBanco.getRange(Linha, 45).setValue(Entrada.getRange('F14').getValue());
      SheetBanco.getRange(Linha, 46).setValue(Entrada.getRange('F15').getValue());
      SheetBanco.getRange(Linha, 47).setValue(Entrada.getRange('F16').getValue());
      SheetBanco.getRange(Linha, 48).setValue(Entrada.getRange('F17').getValue());
      SheetBanco.getRange(Linha, 49).setValue(Entrada.getRange('H2').getValue());
      SheetBanco.getRange(Linha, 50).setValue(Entrada.getRange('H3').getValue());
      SheetBanco.getRange(Linha, 51).setValue(Entrada.getRange('H4').getValue());
      SheetBanco.getRange(Linha, 52).setValue(Entrada.getRange('H5').getValue());
      SheetBanco.getRange(Linha, 53).setValue(Entrada.getRange('H6').getValue());
      SheetBanco.getRange(Linha, 54).setValue(Entrada.getRange('H7').getValue());
      SheetBanco.getRange(Linha, 55).setValue(Entrada.getRange('H8').getValue());
      SheetBanco.getRange(Linha, 56).setValue(Entrada.getRange('H9').getValue());
      SheetBanco.getRange(Linha, 57).setValue(Entrada.getRange('H10').getValue());
      SheetBanco.getRange(Linha, 58).setValue(Entrada.getRange('H11').getValue());
      SheetBanco.getRange(Linha, 59).setValue(Entrada.getRange('H12').getValue());
      SheetBanco.getRange(Linha, 60).setValue(Entrada.getRange('H13').getValue());
      SheetBanco.getRange(Linha, 61).setValue(Entrada.getRange('H14').getValue());
      SheetBanco.getRange(Linha, 62).setValue(Entrada.getRange('H16').getValue());
      SheetBanco.getRange(Linha, 63).setValue(Entrada.getRange('B19').getValue());
      SheetBanco.getRange(Linha, 64).setValue(Entrada.getRange('B20').getValue());
      SheetBanco.getRange(Linha, 65).setValue(Entrada.getRange('B21').getValue());
      SheetBanco.getRange(Linha, 66).setValue(Entrada.getRange('B22').getValue());
      SheetBanco.getRange(Linha, 67).setValue(Entrada.getRange('B23').getValue());
      SheetBanco.getRange(Linha, 68).setValue(Entrada.getRange('B24').getValue());
      SheetBanco.getRange(Linha, 69).setValue(Entrada.getRange('B25').getValue());
      SheetBanco.getRange(Linha, 70).setValue(Entrada.getRange('B26').getValue());
      SheetBanco.getRange(Linha, 71).setValue(Entrada.getRange('B27').getValue());
      SheetBanco.getRange(Linha, 72).setValue(Entrada.getRange('B28').getValue());
      SheetBanco.getRange(Linha, 73).setValue(Entrada.getRange('B29').getValue());
      SheetBanco.getRange(Linha, 74).setValue(Entrada.getRange('B30').getValue());
      SheetBanco.getRange(Linha, 75).setValue(Entrada.getRange('B31').getValue());
      SheetBanco.getRange(Linha, 76).setValue(Entrada.getRange('B32').getValue());
      SheetBanco.getRange(Linha, 77).setValue(Entrada.getRange('B33').getValue());
      SheetBanco.getRange(Linha, 78).setValue(Entrada.getRange('B34').getValue());
      SheetBanco.getRange(Linha, 79).setValue(Entrada.getRange('B35').getValue());
      SheetBanco.getRange(Linha, 80).setValue(Entrada.getRange('B36').getValue());
      SheetBanco.getRange(Linha, 81).setValue(Entrada.getRange('B37').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('B38').getValue());
      SheetBanco.getRange(Linha, 82).setValue(Entrada.getRange('B39').getValue());
      SheetBanco.getRange(Linha, 83).setValue(Entrada.getRange('B40').getValue());
      SheetBanco.getRange(Linha, 84).setValue(Entrada.getRange('B41').getValue());
      SheetBanco.getRange(Linha, 85).setValue(Entrada.getRange('D19').getValue());
      SheetBanco.getRange(Linha, 86).setValue(Entrada.getRange('D20').getValue());
      SheetBanco.getRange(Linha, 87).setValue(Entrada.getRange('D21').getValue());
      SheetBanco.getRange(Linha, 88).setValue(Entrada.getRange('D22').getValue());
      SheetBanco.getRange(Linha, 89).setValue(Entrada.getRange('D23').getValue());
      SheetBanco.getRange(Linha, 90).setValue(Entrada.getRange('D24').getValue());
      SheetBanco.getRange(Linha, 91).setValue(Entrada.getRange('D25').getValue());
      SheetBanco.getRange(Linha, 92).setValue(Entrada.getRange('D26').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('D27').getValue());
      SheetBanco.getRange(Linha, 93).setValue(Entrada.getRange('D28').getValue());
      SheetBanco.getRange(Linha, 94).setValue(Entrada.getRange('D29').getValue());
      SheetBanco.getRange(Linha, 95).setValue(Entrada.getRange('D30').getValue());
      SheetBanco.getRange(Linha, 96).setValue(Entrada.getRange('D31').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('D32').getValue());
      SheetBanco.getRange(Linha, 97).setValue(Entrada.getRange('D33').getValue());
      SheetBanco.getRange(Linha, 98).setValue(Entrada.getRange('D34').getValue());
      SheetBanco.getRange(Linha, 99).setValue(Entrada.getRange('D35').getValue());
      SheetBanco.getRange(Linha, 100).setValue(Entrada.getRange('D36').getValue());
      SheetBanco.getRange(Linha, 101).setValue(Entrada.getRange('D37').getValue());
      SheetBanco.getRange(Linha, 102).setValue(Entrada.getRange('D38').getValue());
      SheetBanco.getRange(Linha, 103).setValue(Entrada.getRange('D39').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('D40').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('D41').getValue());
      SheetBanco.getRange(Linha, 104).setValue(Entrada.getRange('F19').getValue());
      SheetBanco.getRange(Linha, 105).setValue(Entrada.getRange('F20').getValue());
      SheetBanco.getRange(Linha, 106).setValue(Entrada.getRange('F21').getValue());
      SheetBanco.getRange(Linha, 107).setValue(Entrada.getRange('F22').getValue());
      SheetBanco.getRange(Linha, 108).setValue(Entrada.getRange('F23').getValue());
      SheetBanco.getRange(Linha, 109).setValue(Entrada.getRange('F24').getValue());
      SheetBanco.getRange(Linha, 110).setValue(Entrada.getRange('F25').getValue());
      SheetBanco.getRange(Linha, 111).setValue(Entrada.getRange('F26').getValue());
      SheetBanco.getRange(Linha, 112).setValue(Entrada.getRange('F27').getValue());
      SheetBanco.getRange(Linha, 113).setValue(Entrada.getRange('F28').getValue());
      SheetBanco.getRange(Linha, 114).setValue(Entrada.getRange('F29').getValue());
      SheetBanco.getRange(Linha, 115).setValue(Entrada.getRange('F30').getValue());
      SheetBanco.getRange(Linha, 116).setValue(Entrada.getRange('F31').getValue());
      SheetBanco.getRange(Linha, 117).setValue(Entrada.getRange('F32').getValue());
      SheetBanco.getRange(Linha, 118).setValue(Entrada.getRange('F33').getValue());
      SheetBanco.getRange(Linha, 119).setValue(Entrada.getRange('F34').getValue());
      SheetBanco.getRange(Linha, 120).setValue(Entrada.getRange('F35').getValue());
      SheetBanco.getRange(Linha, 121).setValue(Entrada.getRange('F36').getValue());
      SheetBanco.getRange(Linha, 122).setValue(Entrada.getRange('F37').getValue());
      SheetBanco.getRange(Linha, 123).setValue(Entrada.getRange('F38').getValue());
      SheetBanco.getRange(Linha, 124).setValue(Entrada.getRange('F39').getValue());
      SheetBanco.getRange(Linha, 125).setValue(Entrada.getRange('F40').getValue());
      SheetBanco.getRange(Linha, 126).setValue(Entrada.getRange('F41').getValue());
      SheetBanco.getRange(Linha, 127).setValue(Entrada.getRange('H19').getValue());
      SheetBanco.getRange(Linha, 128).setValue(Entrada.getRange('H20').getValue());
      SheetBanco.getRange(Linha, 129).setValue(Entrada.getRange('H21').getValue());
      SheetBanco.getRange(Linha, 130).setValue(Entrada.getRange('H22').getValue());
      SheetBanco.getRange(Linha, 131).setValue(Entrada.getRange('H23').getValue());
      SheetBanco.getRange(Linha, 132).setValue(Entrada.getRange('H24').getValue());
      SheetBanco.getRange(Linha, 133).setValue(Entrada.getRange('H25').getValue());
      SheetBanco.getRange(Linha, 134).setValue(Entrada.getRange('H26').getValue());
      SheetBanco.getRange(Linha, 135).setValue(Entrada.getRange('H27').getValue());
      SheetBanco.getRange(Linha, 136).setValue(Entrada.getRange('H28').getValue());
      SheetBanco.getRange(Linha, 137).setValue(Entrada.getRange('H29').getValue());
      SheetBanco.getRange(Linha, 138).setValue(Entrada.getRange('H30').getValue());
      SheetBanco.getRange(Linha, 139).setValue(Entrada.getRange('H31').getValue());
      SheetBanco.getRange(Linha, 140).setValue(Entrada.getRange('H32').getValue());
      SheetBanco.getRange(Linha, 141).setValue(Entrada.getRange('H33').getValue());
      SheetBanco.getRange(Linha, 142).setValue(Entrada.getRange('H34').getValue());
      SheetBanco.getRange(Linha, 143).setValue(Entrada.getRange('H35').getValue());
      SheetBanco.getRange(Linha, 144).setValue(Entrada.getRange('H36').getValue());
      SheetBanco.getRange(Linha, 145).setValue(Entrada.getRange('H37').getValue());
      SheetBanco.getRange(Linha, 146).setValue(Entrada.getRange('H38').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('H39').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('H40').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('H41').getValue());
      SheetBanco.getRange(Linha, 147).setValue(Entrada.getRange('B43').getValue());
      SheetBanco.getRange(Linha, 148).setValue(Entrada.getRange('B44').getValue());
      SheetBanco.getRange(Linha, 149).setValue(Entrada.getRange('B45').getValue());
      SheetBanco.getRange(Linha, 150).setValue(Entrada.getRange('B46').getValue());
      SheetBanco.getRange(Linha, 151).setValue(Entrada.getRange('B47').getValue());
      SheetBanco.getRange(Linha, 152).setValue(Entrada.getRange('B48').getValue());
      SheetBanco.getRange(Linha, 153).setValue(Entrada.getRange('B49').getValue());
      SheetBanco.getRange(Linha, 154).setValue(Entrada.getRange('B50').getValue());
      SheetBanco.getRange(Linha, 155).setValue(Entrada.getRange('B51').getValue());
      SheetBanco.getRange(Linha, 156).setValue(Entrada.getRange('B52').getValue());
      SheetBanco.getRange(Linha, 157).setValue(Entrada.getRange('B53').getValue());
      SheetBanco.getRange(Linha, 158).setValue(Entrada.getRange('B54').getValue());
      SheetBanco.getRange(Linha, 159).setValue(Entrada.getRange('B55').getValue());
      SheetBanco.getRange(Linha, 160).setValue(Entrada.getRange('B56').getValue());
      SheetBanco.getRange(Linha, 161).setValue(Entrada.getRange('B57').getValue());
      SheetBanco.getRange(Linha, 162).setValue(Entrada.getRange('B58').getValue());
      SheetBanco.getRange(Linha, 163).setValue(Entrada.getRange('B59').getValue());
      SheetBanco.getRange(Linha, 164).setValue(Entrada.getRange('B60').getValue());
      SheetBanco.getRange(Linha, 165).setValue(Entrada.getRange('B61').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('B62').getValue());
      SheetBanco.getRange(Linha, 166).setValue(Entrada.getRange('B63').getValue());
      SheetBanco.getRange(Linha, 167).setValue(Entrada.getRange('B64').getValue());
      SheetBanco.getRange(Linha, 168).setValue(Entrada.getRange('B65').getValue());
      SheetBanco.getRange(Linha, 169).setValue(Entrada.getRange('B66').getValue());
      SheetBanco.getRange(Linha, 170).setValue(Entrada.getRange('B67').getValue());
      SheetBanco.getRange(Linha, 171).setValue(Entrada.getRange('B68').getValue());
      SheetBanco.getRange(Linha, 172).setValue(Entrada.getRange('B69').getValue());
      SheetBanco.getRange(Linha, 173).setValue(Entrada.getRange('B70').getValue());
      SheetBanco.getRange(Linha, 174).setValue(Entrada.getRange('B71').getValue());
      SheetBanco.getRange(Linha, 175).setValue(Entrada.getRange('B72').getValue());
      SheetBanco.getRange(Linha, 176).setValue(Entrada.getRange('B73').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('B74').getValue());
      SheetBanco.getRange(Linha, 177).setValue(Entrada.getRange('B75').getValue());
      SheetBanco.getRange(Linha, 178).setValue(Entrada.getRange('B76').getValue());
      SheetBanco.getRange(Linha, 179).setValue(Entrada.getRange('B77').getValue());
      SheetBanco.getRange(Linha, 180).setValue(Entrada.getRange('B78').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('B79').getValue());
      SheetBanco.getRange(Linha, 181).setValue(Entrada.getRange('B80').getValue());
      SheetBanco.getRange(Linha, 182).setValue(Entrada.getRange('B81').getValue());
      SheetBanco.getRange(Linha, 183).setValue(Entrada.getRange('B82').getValue());
      SheetBanco.getRange(Linha, 184).setValue(Entrada.getRange('B83').getValue());
      SheetBanco.getRange(Linha, 185).setValue(Entrada.getRange('B84').getValue());
      SheetBanco.getRange(Linha, 186).setValue(Entrada.getRange('B85').getValue());
      SheetBanco.getRange(Linha, 187).setValue(Entrada.getRange('B86').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('B87').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('B88').getValue());
      SheetBanco.getRange(Linha, 188).setValue(Entrada.getRange('D43').getValue());
      SheetBanco.getRange(Linha, 189).setValue(Entrada.getRange('D44').getValue());
      SheetBanco.getRange(Linha, 190).setValue(Entrada.getRange('D45').getValue());
      SheetBanco.getRange(Linha, 191).setValue(Entrada.getRange('D46').getValue());
      SheetBanco.getRange(Linha, 192).setValue(Entrada.getRange('D47').getValue());
      SheetBanco.getRange(Linha, 193).setValue(Entrada.getRange('D48').getValue());
      SheetBanco.getRange(Linha, 194).setValue(Entrada.getRange('D49').getValue());
      SheetBanco.getRange(Linha, 195).setValue(Entrada.getRange('D50').getValue());
      SheetBanco.getRange(Linha, 196).setValue(Entrada.getRange('D51').getValue());
      SheetBanco.getRange(Linha, 197).setValue(Entrada.getRange('D52').getValue());
      SheetBanco.getRange(Linha, 198).setValue(Entrada.getRange('D53').getValue());
      SheetBanco.getRange(Linha, 199).setValue(Entrada.getRange('D54').getValue());
      SheetBanco.getRange(Linha, 200).setValue(Entrada.getRange('D55').getValue());
      SheetBanco.getRange(Linha, 201).setValue(Entrada.getRange('D56').getValue());
      SheetBanco.getRange(Linha, 202).setValue(Entrada.getRange('D57').getValue());
      SheetBanco.getRange(Linha, 203).setValue(Entrada.getRange('D58').getValue());
      SheetBanco.getRange(Linha, 204).setValue(Entrada.getRange('D59').getValue());
      SheetBanco.getRange(Linha, 205).setValue(Entrada.getRange('D60').getValue());
      SheetBanco.getRange(Linha, 206).setValue(Entrada.getRange('D61').getValue());
      SheetBanco.getRange(Linha, 207).setValue(Entrada.getRange('D62').getValue());
      SheetBanco.getRange(Linha, 208).setValue(Entrada.getRange('D63').getValue());
      SheetBanco.getRange(Linha, 209).setValue(Entrada.getRange('D64').getValue());
      SheetBanco.getRange(Linha, 210).setValue(Entrada.getRange('D65').getValue());
      SheetBanco.getRange(Linha, 211).setValue(Entrada.getRange('D66').getValue());
      SheetBanco.getRange(Linha, 212).setValue(Entrada.getRange('D67').getValue());
      SheetBanco.getRange(Linha, 213).setValue(Entrada.getRange('D68').getValue());
      SheetBanco.getRange(Linha, 214).setValue(Entrada.getRange('D69').getValue());
      SheetBanco.getRange(Linha, 215).setValue(Entrada.getRange('D70').getValue());
      SheetBanco.getRange(Linha, 216).setValue(Entrada.getRange('D71').getValue());
      SheetBanco.getRange(Linha, 217).setValue(Entrada.getRange('D72').getValue());
      SheetBanco.getRange(Linha, 218).setValue(Entrada.getRange('D73').getValue());
      SheetBanco.getRange(Linha, 219).setValue(Entrada.getRange('D74').getValue());
      SheetBanco.getRange(Linha, 220).setValue(Entrada.getRange('D75').getValue());
      SheetBanco.getRange(Linha, 221).setValue(Entrada.getRange('D76').getValue());
      SheetBanco.getRange(Linha, 222).setValue(Entrada.getRange('D77').getValue());
      SheetBanco.getRange(Linha, 223).setValue(Entrada.getRange('D78').getValue());
      SheetBanco.getRange(Linha, 224).setValue(Entrada.getRange('D79').getValue());
      SheetBanco.getRange(Linha, 225).setValue(Entrada.getRange('D80').getValue());
      SheetBanco.getRange(Linha, 226).setValue(Entrada.getRange('D81').getValue());
      SheetBanco.getRange(Linha, 227).setValue(Entrada.getRange('D82').getValue());
      SheetBanco.getRange(Linha, 228).setValue(Entrada.getRange('D83').getValue());
      SheetBanco.getRange(Linha, 229).setValue(Entrada.getRange('D84').getValue());
      SheetBanco.getRange(Linha, 230).setValue(Entrada.getRange('D85').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('D86').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('D87').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('D88').getValue());
      SheetBanco.getRange(Linha, 231).setValue(Entrada.getRange('F43').getValue());
      SheetBanco.getRange(Linha, 232).setValue(Entrada.getRange('F44').getValue());
      SheetBanco.getRange(Linha, 233).setValue(Entrada.getRange('F45').getValue());
      SheetBanco.getRange(Linha, 234).setValue(Entrada.getRange('F46').getValue());
      SheetBanco.getRange(Linha, 235).setValue(Entrada.getRange('F47').getValue());
      SheetBanco.getRange(Linha, 236).setValue(Entrada.getRange('F48').getValue());
      SheetBanco.getRange(Linha, 237).setValue(Entrada.getRange('F49').getValue());
      SheetBanco.getRange(Linha, 238).setValue(Entrada.getRange('F50').getValue());
      SheetBanco.getRange(Linha, 239).setValue(Entrada.getRange('F51').getValue());
      SheetBanco.getRange(Linha, 240).setValue(Entrada.getRange('F52').getValue());
      SheetBanco.getRange(Linha, 241).setValue(Entrada.getRange('F53').getValue());
      SheetBanco.getRange(Linha, 242).setValue(Entrada.getRange('F54').getValue());
      SheetBanco.getRange(Linha, 243).setValue(Entrada.getRange('F55').getValue());
      SheetBanco.getRange(Linha, 244).setValue(Entrada.getRange('F56').getValue());
      SheetBanco.getRange(Linha, 245).setValue(Entrada.getRange('F57').getValue());
      SheetBanco.getRange(Linha, 246).setValue(Entrada.getRange('F58').getValue());
      SheetBanco.getRange(Linha, 247).setValue(Entrada.getRange('F59').getValue());
      SheetBanco.getRange(Linha, 248).setValue(Entrada.getRange('F60').getValue());
      SheetBanco.getRange(Linha, 249).setValue(Entrada.getRange('F61').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('F62').getValue());
      SheetBanco.getRange(Linha, 250).setValue(Entrada.getRange('F63').getValue());
      SheetBanco.getRange(Linha, 251).setValue(Entrada.getRange('F64').getValue());
      SheetBanco.getRange(Linha, 252).setValue(Entrada.getRange('F65').getValue());
      SheetBanco.getRange(Linha, 253).setValue(Entrada.getRange('F66').getValue());
      SheetBanco.getRange(Linha, 254).setValue(Entrada.getRange('F67').getValue());
      SheetBanco.getRange(Linha, 255).setValue(Entrada.getRange('F68').getValue());
      SheetBanco.getRange(Linha, 256).setValue(Entrada.getRange('F69').getValue());
      SheetBanco.getRange(Linha, 257).setValue(Entrada.getRange('F70').getValue());
      SheetBanco.getRange(Linha, 258).setValue(Entrada.getRange('F71').getValue());
      SheetBanco.getRange(Linha, 259).setValue(Entrada.getRange('F72').getValue());
      SheetBanco.getRange(Linha, 260).setValue(Entrada.getRange('F73').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('F74').getValue());
      SheetBanco.getRange(Linha, 261).setValue(Entrada.getRange('F75').getValue());
      SheetBanco.getRange(Linha, 262).setValue(Entrada.getRange('F76').getValue());
      SheetBanco.getRange(Linha, 263).setValue(Entrada.getRange('F77').getValue());
      SheetBanco.getRange(Linha, 264).setValue(Entrada.getRange('F78').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('F79').getValue());
      SheetBanco.getRange(Linha, 265).setValue(Entrada.getRange('F80').getValue());
      SheetBanco.getRange(Linha, 266).setValue(Entrada.getRange('F81').getValue());
      SheetBanco.getRange(Linha, 267).setValue(Entrada.getRange('F82').getValue());
      SheetBanco.getRange(Linha, 268).setValue(Entrada.getRange('F83').getValue());
      SheetBanco.getRange(Linha, 269).setValue(Entrada.getRange('F84').getValue());
      SheetBanco.getRange(Linha, 270).setValue(Entrada.getRange('F85').getValue());
      SheetBanco.getRange(Linha, 271).setValue(Entrada.getRange('F86').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('F87').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('F88').getValue());
      SheetBanco.getRange(Linha, 272).setValue(Entrada.getRange('H43').getValue());
      SheetBanco.getRange(Linha, 273).setValue(Entrada.getRange('H44').getValue());
      SheetBanco.getRange(Linha, 274).setValue(Entrada.getRange('H45').getValue());
      SheetBanco.getRange(Linha, 275).setValue(Entrada.getRange('H46').getValue());
      SheetBanco.getRange(Linha, 276).setValue(Entrada.getRange('H47').getValue());
      SheetBanco.getRange(Linha, 277).setValue(Entrada.getRange('H48').getValue());
      SheetBanco.getRange(Linha, 278).setValue(Entrada.getRange('H49').getValue());
      SheetBanco.getRange(Linha, 279).setValue(Entrada.getRange('H50').getValue());
      SheetBanco.getRange(Linha, 280).setValue(Entrada.getRange('H51').getValue());
      SheetBanco.getRange(Linha, 281).setValue(Entrada.getRange('H52').getValue());
      SheetBanco.getRange(Linha, 282).setValue(Entrada.getRange('H53').getValue());
      SheetBanco.getRange(Linha, 283).setValue(Entrada.getRange('H54').getValue());
      SheetBanco.getRange(Linha, 284).setValue(Entrada.getRange('H55').getValue());
      SheetBanco.getRange(Linha, 285).setValue(Entrada.getRange('H56').getValue());
      SheetBanco.getRange(Linha, 286).setValue(Entrada.getRange('H57').getValue());
      SheetBanco.getRange(Linha, 287).setValue(Entrada.getRange('H58').getValue());
      SheetBanco.getRange(Linha, 288).setValue(Entrada.getRange('H59').getValue());
      SheetBanco.getRange(Linha, 289).setValue(Entrada.getRange('H60').getValue());
      SheetBanco.getRange(Linha, 290).setValue(Entrada.getRange('H61').getValue());
      SheetBanco.getRange(Linha, 291).setValue(Entrada.getRange('H62').getValue());
      SheetBanco.getRange(Linha, 292).setValue(Entrada.getRange('H63').getValue());
      SheetBanco.getRange(Linha, 293).setValue(Entrada.getRange('H64').getValue());
      SheetBanco.getRange(Linha, 294).setValue(Entrada.getRange('H65').getValue());
      SheetBanco.getRange(Linha, 295).setValue(Entrada.getRange('H66').getValue());
      SheetBanco.getRange(Linha, 296).setValue(Entrada.getRange('H67').getValue());
      SheetBanco.getRange(Linha, 297).setValue(Entrada.getRange('H68').getValue());
      SheetBanco.getRange(Linha, 298).setValue(Entrada.getRange('H69').getValue());
      SheetBanco.getRange(Linha, 299).setValue(Entrada.getRange('H70').getValue());
      SheetBanco.getRange(Linha, 300).setValue(Entrada.getRange('H71').getValue());
      SheetBanco.getRange(Linha, 301).setValue(Entrada.getRange('H72').getValue());
      SheetBanco.getRange(Linha, 302).setValue(Entrada.getRange('H73').getValue());
      SheetBanco.getRange(Linha, 303).setValue(Entrada.getRange('H74').getValue());
      SheetBanco.getRange(Linha, 304).setValue(Entrada.getRange('H75').getValue());
      SheetBanco.getRange(Linha, 305).setValue(Entrada.getRange('H76').getValue());
      SheetBanco.getRange(Linha, 306).setValue(Entrada.getRange('H77').getValue());
      SheetBanco.getRange(Linha, 307).setValue(Entrada.getRange('H78').getValue());
      SheetBanco.getRange(Linha, 308).setValue(Entrada.getRange('H79').getValue());
      SheetBanco.getRange(Linha, 309).setValue(Entrada.getRange('H80').getValue());
      SheetBanco.getRange(Linha, 310).setValue(Entrada.getRange('H81').getValue());
      SheetBanco.getRange(Linha, 311).setValue(Entrada.getRange('H82').getValue());
      SheetBanco.getRange(Linha, 312).setValue(Entrada.getRange('H83').getValue());
      SheetBanco.getRange(Linha, 313).setValue(Entrada.getRange('H84').getValue());
      SheetBanco.getRange(Linha, 314).setValue(Entrada.getRange('H85').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('H86').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('H87').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('H88').getValue());
      SheetBanco.getRange(Linha, 315).setValue(Entrada.getRange('B90').getValue());
      SheetBanco.getRange(Linha, 316).setValue(Entrada.getRange('B91').getValue());
      SheetBanco.getRange(Linha, 317).setValue(Entrada.getRange('B92').getValue());
      SheetBanco.getRange(Linha, 318).setValue(Entrada.getRange('B93').getValue());
      SheetBanco.getRange(Linha, 319).setValue(Entrada.getRange('B94').getValue());
      SheetBanco.getRange(Linha, 320).setValue(Entrada.getRange('B95').getValue());
      SheetBanco.getRange(Linha, 321).setValue(Entrada.getRange('B96').getValue());
      SheetBanco.getRange(Linha, 322).setValue(Entrada.getRange('B97').getValue());
      SheetBanco.getRange(Linha, 323).setValue(Entrada.getRange('B98').getValue());
      SheetBanco.getRange(Linha, 324).setValue(Entrada.getRange('B99').getValue());
      SheetBanco.getRange(Linha, 325).setValue(Entrada.getRange('B100').getValue());
      SheetBanco.getRange(Linha, 326).setValue(Entrada.getRange('B101').getValue());
      SheetBanco.getRange(Linha, 327).setValue(Entrada.getRange('B102').getValue());
      SheetBanco.getRange(Linha, 328).setValue(Entrada.getRange('B103').getValue());
      SheetBanco.getRange(Linha, 329).setValue(Entrada.getRange('B104').getValue());
      SheetBanco.getRange(Linha, 330).setValue(Entrada.getRange('B105').getValue());
      SheetBanco.getRange(Linha, 331).setValue(Entrada.getRange('B106').getValue());
      SheetBanco.getRange(Linha, 332).setValue(Entrada.getRange('B107').getValue());
      SheetBanco.getRange(Linha, 333).setValue(Entrada.getRange('B108').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('B109').getValue());
      SheetBanco.getRange(Linha, 334).setValue(Entrada.getRange('B110').getValue());
      SheetBanco.getRange(Linha, 335).setValue(Entrada.getRange('B111').getValue());
      SheetBanco.getRange(Linha, 336).setValue(Entrada.getRange('B112').getValue());
      SheetBanco.getRange(Linha, 337).setValue(Entrada.getRange('B113').getValue());
      SheetBanco.getRange(Linha, 338).setValue(Entrada.getRange('B114').getValue());
      SheetBanco.getRange(Linha, 339).setValue(Entrada.getRange('B115').getValue());
      SheetBanco.getRange(Linha, 340).setValue(Entrada.getRange('B116').getValue());
      SheetBanco.getRange(Linha, 341).setValue(Entrada.getRange('B117').getValue());
      SheetBanco.getRange(Linha, 342).setValue(Entrada.getRange('B118').getValue());
      SheetBanco.getRange(Linha, 343).setValue(Entrada.getRange('B119').getValue());
      SheetBanco.getRange(Linha, 344).setValue(Entrada.getRange('B120').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('B121').getValue());
      SheetBanco.getRange(Linha, 345).setValue(Entrada.getRange('B122').getValue());
      SheetBanco.getRange(Linha, 346).setValue(Entrada.getRange('B123').getValue());
      SheetBanco.getRange(Linha, 347).setValue(Entrada.getRange('B124').getValue());
      SheetBanco.getRange(Linha, 348).setValue(Entrada.getRange('B125').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('B126').getValue());
      SheetBanco.getRange(Linha, 349).setValue(Entrada.getRange('B127').getValue());
      SheetBanco.getRange(Linha, 350).setValue(Entrada.getRange('B128').getValue());
      SheetBanco.getRange(Linha, 351).setValue(Entrada.getRange('B129').getValue());
      SheetBanco.getRange(Linha, 352).setValue(Entrada.getRange('B130').getValue());
      SheetBanco.getRange(Linha, 353).setValue(Entrada.getRange('B131').getValue());
      SheetBanco.getRange(Linha, 354).setValue(Entrada.getRange('B132').getValue());
      SheetBanco.getRange(Linha, 355).setValue(Entrada.getRange('B133').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('B134').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('B135').getValue());
      SheetBanco.getRange(Linha, 356).setValue(Entrada.getRange('D90').getValue());
      SheetBanco.getRange(Linha, 357).setValue(Entrada.getRange('D91').getValue());
      SheetBanco.getRange(Linha, 358).setValue(Entrada.getRange('D92').getValue());
      SheetBanco.getRange(Linha, 359).setValue(Entrada.getRange('D93').getValue());
      SheetBanco.getRange(Linha, 360).setValue(Entrada.getRange('D94').getValue());
      SheetBanco.getRange(Linha, 361).setValue(Entrada.getRange('D95').getValue());
      SheetBanco.getRange(Linha, 362).setValue(Entrada.getRange('D96').getValue());
      SheetBanco.getRange(Linha, 363).setValue(Entrada.getRange('D97').getValue());
      SheetBanco.getRange(Linha, 364).setValue(Entrada.getRange('D98').getValue());
      SheetBanco.getRange(Linha, 365).setValue(Entrada.getRange('D99').getValue());
      SheetBanco.getRange(Linha, 366).setValue(Entrada.getRange('D100').getValue());
      SheetBanco.getRange(Linha, 367).setValue(Entrada.getRange('D101').getValue());
      SheetBanco.getRange(Linha, 368).setValue(Entrada.getRange('D102').getValue());
      SheetBanco.getRange(Linha, 369).setValue(Entrada.getRange('D103').getValue());
      SheetBanco.getRange(Linha, 370).setValue(Entrada.getRange('D104').getValue());
      SheetBanco.getRange(Linha, 371).setValue(Entrada.getRange('D105').getValue());
      SheetBanco.getRange(Linha, 372).setValue(Entrada.getRange('D106').getValue());
      SheetBanco.getRange(Linha, 373).setValue(Entrada.getRange('D107').getValue());
      SheetBanco.getRange(Linha, 374).setValue(Entrada.getRange('D108').getValue());
      SheetBanco.getRange(Linha, 375).setValue(Entrada.getRange('D109').getValue());
      SheetBanco.getRange(Linha, 376).setValue(Entrada.getRange('D110').getValue());
      SheetBanco.getRange(Linha, 377).setValue(Entrada.getRange('D111').getValue());
      SheetBanco.getRange(Linha, 378).setValue(Entrada.getRange('D112').getValue());
      SheetBanco.getRange(Linha, 379).setValue(Entrada.getRange('D113').getValue());
      SheetBanco.getRange(Linha, 380).setValue(Entrada.getRange('D114').getValue());
      SheetBanco.getRange(Linha, 381).setValue(Entrada.getRange('D115').getValue());
      SheetBanco.getRange(Linha, 382).setValue(Entrada.getRange('D116').getValue());
      SheetBanco.getRange(Linha, 383).setValue(Entrada.getRange('D117').getValue());
      SheetBanco.getRange(Linha, 384).setValue(Entrada.getRange('D118').getValue());
      SheetBanco.getRange(Linha, 385).setValue(Entrada.getRange('D119').getValue());
      SheetBanco.getRange(Linha, 386).setValue(Entrada.getRange('D120').getValue());
      SheetBanco.getRange(Linha, 387).setValue(Entrada.getRange('D121').getValue());
      SheetBanco.getRange(Linha, 388).setValue(Entrada.getRange('D122').getValue());
      SheetBanco.getRange(Linha, 389).setValue(Entrada.getRange('D123').getValue());
      SheetBanco.getRange(Linha, 390).setValue(Entrada.getRange('D124').getValue());
      SheetBanco.getRange(Linha, 391).setValue(Entrada.getRange('D125').getValue());
      SheetBanco.getRange(Linha, 392).setValue(Entrada.getRange('D126').getValue());
      SheetBanco.getRange(Linha, 393).setValue(Entrada.getRange('D127').getValue());
      SheetBanco.getRange(Linha, 394).setValue(Entrada.getRange('D128').getValue());
      SheetBanco.getRange(Linha, 395).setValue(Entrada.getRange('D129').getValue());
      SheetBanco.getRange(Linha, 396).setValue(Entrada.getRange('D130').getValue());
      SheetBanco.getRange(Linha, 397).setValue(Entrada.getRange('D131').getValue());
      SheetBanco.getRange(Linha, 398).setValue(Entrada.getRange('D132').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('D133').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('D134').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('D135').getValue());
      SheetBanco.getRange(Linha, 399).setValue(Entrada.getRange('F90').getValue());
      SheetBanco.getRange(Linha, 400).setValue(Entrada.getRange('F91').getValue());
      SheetBanco.getRange(Linha, 401).setValue(Entrada.getRange('F92').getValue());
      SheetBanco.getRange(Linha, 402).setValue(Entrada.getRange('F93').getValue());
      SheetBanco.getRange(Linha, 403).setValue(Entrada.getRange('F94').getValue());
      SheetBanco.getRange(Linha, 404).setValue(Entrada.getRange('F95').getValue());
      SheetBanco.getRange(Linha, 405).setValue(Entrada.getRange('F96').getValue());
      SheetBanco.getRange(Linha, 406).setValue(Entrada.getRange('F97').getValue());
      SheetBanco.getRange(Linha, 407).setValue(Entrada.getRange('F98').getValue());
      SheetBanco.getRange(Linha, 408).setValue(Entrada.getRange('F99').getValue());
      SheetBanco.getRange(Linha, 409).setValue(Entrada.getRange('F100').getValue());
      SheetBanco.getRange(Linha, 410).setValue(Entrada.getRange('F101').getValue());
      SheetBanco.getRange(Linha, 411).setValue(Entrada.getRange('F102').getValue());
      SheetBanco.getRange(Linha, 412).setValue(Entrada.getRange('F103').getValue());
      SheetBanco.getRange(Linha, 413).setValue(Entrada.getRange('F104').getValue());
      SheetBanco.getRange(Linha, 414).setValue(Entrada.getRange('F105').getValue());
      SheetBanco.getRange(Linha, 415).setValue(Entrada.getRange('F106').getValue());
      SheetBanco.getRange(Linha, 416).setValue(Entrada.getRange('F107').getValue());
      SheetBanco.getRange(Linha, 417).setValue(Entrada.getRange('F108').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('F109').getValue());
      SheetBanco.getRange(Linha, 418).setValue(Entrada.getRange('F110').getValue());
      SheetBanco.getRange(Linha, 419).setValue(Entrada.getRange('F111').getValue());
      SheetBanco.getRange(Linha, 420).setValue(Entrada.getRange('F112').getValue());
      SheetBanco.getRange(Linha, 421).setValue(Entrada.getRange('F113').getValue());
      SheetBanco.getRange(Linha, 422).setValue(Entrada.getRange('F114').getValue());
      SheetBanco.getRange(Linha, 423).setValue(Entrada.getRange('F115').getValue());
      SheetBanco.getRange(Linha, 424).setValue(Entrada.getRange('F116').getValue());
      SheetBanco.getRange(Linha, 425).setValue(Entrada.getRange('F117').getValue());
      SheetBanco.getRange(Linha, 426).setValue(Entrada.getRange('F118').getValue());
      SheetBanco.getRange(Linha, 427).setValue(Entrada.getRange('F119').getValue());
      SheetBanco.getRange(Linha, 428).setValue(Entrada.getRange('F120').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('F121').getValue());
      SheetBanco.getRange(Linha, 429).setValue(Entrada.getRange('F122').getValue());
      SheetBanco.getRange(Linha, 430).setValue(Entrada.getRange('F123').getValue());
      SheetBanco.getRange(Linha, 431).setValue(Entrada.getRange('F124').getValue());
      SheetBanco.getRange(Linha, 432).setValue(Entrada.getRange('F125').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('F126').getValue());
      SheetBanco.getRange(Linha, 433).setValue(Entrada.getRange('F127').getValue());
      SheetBanco.getRange(Linha, 434).setValue(Entrada.getRange('F128').getValue());
      SheetBanco.getRange(Linha, 435).setValue(Entrada.getRange('F129').getValue());
      SheetBanco.getRange(Linha, 436).setValue(Entrada.getRange('F130').getValue());
      SheetBanco.getRange(Linha, 437).setValue(Entrada.getRange('F131').getValue());
      SheetBanco.getRange(Linha, 438).setValue(Entrada.getRange('F132').getValue());
      SheetBanco.getRange(Linha, 439).setValue(Entrada.getRange('F133').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('F134').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('F135').getValue());
      SheetBanco.getRange(Linha, 440).setValue(Entrada.getRange('H90').getValue());
      SheetBanco.getRange(Linha, 441).setValue(Entrada.getRange('H91').getValue());
      SheetBanco.getRange(Linha, 442).setValue(Entrada.getRange('H92').getValue());
      SheetBanco.getRange(Linha, 443).setValue(Entrada.getRange('H93').getValue());
      SheetBanco.getRange(Linha, 444).setValue(Entrada.getRange('H94').getValue());
      SheetBanco.getRange(Linha, 445).setValue(Entrada.getRange('H95').getValue());
      SheetBanco.getRange(Linha, 446).setValue(Entrada.getRange('H96').getValue());
      SheetBanco.getRange(Linha, 447).setValue(Entrada.getRange('H97').getValue());
      SheetBanco.getRange(Linha, 448).setValue(Entrada.getRange('H98').getValue());
      SheetBanco.getRange(Linha, 449).setValue(Entrada.getRange('H99').getValue());
      SheetBanco.getRange(Linha, 450).setValue(Entrada.getRange('H100').getValue());
      SheetBanco.getRange(Linha, 451).setValue(Entrada.getRange('H101').getValue());
      SheetBanco.getRange(Linha, 452).setValue(Entrada.getRange('H102').getValue());
      SheetBanco.getRange(Linha, 453).setValue(Entrada.getRange('H103').getValue());
      SheetBanco.getRange(Linha, 454).setValue(Entrada.getRange('H104').getValue());
      SheetBanco.getRange(Linha, 455).setValue(Entrada.getRange('H105').getValue());
      SheetBanco.getRange(Linha, 456).setValue(Entrada.getRange('H106').getValue());
      SheetBanco.getRange(Linha, 457).setValue(Entrada.getRange('H107').getValue());
      SheetBanco.getRange(Linha, 458).setValue(Entrada.getRange('H108').getValue());
      SheetBanco.getRange(Linha, 459).setValue(Entrada.getRange('H109').getValue());
      SheetBanco.getRange(Linha, 460).setValue(Entrada.getRange('H110').getValue());
      SheetBanco.getRange(Linha, 461).setValue(Entrada.getRange('H111').getValue());
      SheetBanco.getRange(Linha, 462).setValue(Entrada.getRange('H112').getValue());
      SheetBanco.getRange(Linha, 463).setValue(Entrada.getRange('H113').getValue());
      SheetBanco.getRange(Linha, 464).setValue(Entrada.getRange('H114').getValue());
      SheetBanco.getRange(Linha, 465).setValue(Entrada.getRange('H115').getValue());
      SheetBanco.getRange(Linha, 466).setValue(Entrada.getRange('H116').getValue());
      SheetBanco.getRange(Linha, 467).setValue(Entrada.getRange('H117').getValue());
      SheetBanco.getRange(Linha, 468).setValue(Entrada.getRange('H118').getValue());
      SheetBanco.getRange(Linha, 469).setValue(Entrada.getRange('H119').getValue());
      SheetBanco.getRange(Linha, 470).setValue(Entrada.getRange('H120').getValue());
      SheetBanco.getRange(Linha, 471).setValue(Entrada.getRange('H121').getValue());
      SheetBanco.getRange(Linha, 472).setValue(Entrada.getRange('H122').getValue());
      SheetBanco.getRange(Linha, 473).setValue(Entrada.getRange('H123').getValue());
      SheetBanco.getRange(Linha, 474).setValue(Entrada.getRange('H124').getValue());
      SheetBanco.getRange(Linha, 475).setValue(Entrada.getRange('H125').getValue());
      SheetBanco.getRange(Linha, 476).setValue(Entrada.getRange('H126').getValue());
      SheetBanco.getRange(Linha, 477).setValue(Entrada.getRange('H127').getValue());
      SheetBanco.getRange(Linha, 478).setValue(Entrada.getRange('H128').getValue());
      SheetBanco.getRange(Linha, 479).setValue(Entrada.getRange('H129').getValue());
      SheetBanco.getRange(Linha, 480).setValue(Entrada.getRange('H130').getValue());
      SheetBanco.getRange(Linha, 481).setValue(Entrada.getRange('H131').getValue());
      SheetBanco.getRange(Linha, 482).setValue(Entrada.getRange('H132').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('H133').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('H134').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('H135').getValue());
      SheetBanco.getRange(Linha, 483).setValue(Entrada.getRange('B137').getValue());
      SheetBanco.getRange(Linha, 484).setValue(Entrada.getRange('B138').getValue());
      SheetBanco.getRange(Linha, 485).setValue(Entrada.getRange('B139').getValue());
      SheetBanco.getRange(Linha, 486).setValue(Entrada.getRange('B140').getValue());
      SheetBanco.getRange(Linha, 487).setValue(Entrada.getRange('B141').getValue());
      SheetBanco.getRange(Linha, 488).setValue(Entrada.getRange('B142').getValue());
      SheetBanco.getRange(Linha, 489).setValue(Entrada.getRange('B143').getValue());
      SheetBanco.getRange(Linha, 490).setValue(Entrada.getRange('B144').getValue());
      SheetBanco.getRange(Linha, 491).setValue(Entrada.getRange('B145').getValue());
      SheetBanco.getRange(Linha, 492).setValue(Entrada.getRange('B146').getValue());
      SheetBanco.getRange(Linha, 493).setValue(Entrada.getRange('B147').getValue());
      SheetBanco.getRange(Linha, 494).setValue(Entrada.getRange('B148').getValue());
      SheetBanco.getRange(Linha, 495).setValue(Entrada.getRange('B149').getValue());
      SheetBanco.getRange(Linha, 496).setValue(Entrada.getRange('B150').getValue());
      SheetBanco.getRange(Linha, 497).setValue(Entrada.getRange('B151').getValue());
      SheetBanco.getRange(Linha, 498).setValue(Entrada.getRange('B152').getValue());
      SheetBanco.getRange(Linha, 499).setValue(Entrada.getRange('B153').getValue());
      SheetBanco.getRange(Linha, 500).setValue(Entrada.getRange('B154').getValue());
      SheetBanco.getRange(Linha, 501).setValue(Entrada.getRange('B155').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('B156').getValue());
      SheetBanco.getRange(Linha, 502).setValue(Entrada.getRange('B157').getValue());
      SheetBanco.getRange(Linha, 503).setValue(Entrada.getRange('B158').getValue());
      SheetBanco.getRange(Linha, 504).setValue(Entrada.getRange('B159').getValue());
      SheetBanco.getRange(Linha, 505).setValue(Entrada.getRange('B160').getValue());
      SheetBanco.getRange(Linha, 506).setValue(Entrada.getRange('B161').getValue());
      SheetBanco.getRange(Linha, 507).setValue(Entrada.getRange('B162').getValue());
      SheetBanco.getRange(Linha, 508).setValue(Entrada.getRange('B163').getValue());
      SheetBanco.getRange(Linha, 509).setValue(Entrada.getRange('B164').getValue());
      SheetBanco.getRange(Linha, 510).setValue(Entrada.getRange('B165').getValue());
      SheetBanco.getRange(Linha, 511).setValue(Entrada.getRange('B166').getValue());
      SheetBanco.getRange(Linha, 512).setValue(Entrada.getRange('B167').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('B168').getValue());
      SheetBanco.getRange(Linha, 513).setValue(Entrada.getRange('B169').getValue());
      SheetBanco.getRange(Linha, 514).setValue(Entrada.getRange('B170').getValue());
      SheetBanco.getRange(Linha, 515).setValue(Entrada.getRange('B171').getValue());
      SheetBanco.getRange(Linha, 516).setValue(Entrada.getRange('B172').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('B173').getValue());
      SheetBanco.getRange(Linha, 517).setValue(Entrada.getRange('B174').getValue());
      SheetBanco.getRange(Linha, 518).setValue(Entrada.getRange('B175').getValue());
      SheetBanco.getRange(Linha, 519).setValue(Entrada.getRange('B176').getValue());
      SheetBanco.getRange(Linha, 520).setValue(Entrada.getRange('B177').getValue());
      SheetBanco.getRange(Linha, 521).setValue(Entrada.getRange('B178').getValue());
      SheetBanco.getRange(Linha, 522).setValue(Entrada.getRange('B179').getValue());
      SheetBanco.getRange(Linha, 523).setValue(Entrada.getRange('B180').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('B181').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('B182').getValue());
      SheetBanco.getRange(Linha, 524).setValue(Entrada.getRange('D137').getValue());
      SheetBanco.getRange(Linha, 525).setValue(Entrada.getRange('D138').getValue());
      SheetBanco.getRange(Linha, 526).setValue(Entrada.getRange('D139').getValue());
      SheetBanco.getRange(Linha, 527).setValue(Entrada.getRange('D140').getValue());
      SheetBanco.getRange(Linha, 528).setValue(Entrada.getRange('D141').getValue());
      SheetBanco.getRange(Linha, 529).setValue(Entrada.getRange('D142').getValue());
      SheetBanco.getRange(Linha, 530).setValue(Entrada.getRange('D143').getValue());
      SheetBanco.getRange(Linha, 531).setValue(Entrada.getRange('D144').getValue());
      SheetBanco.getRange(Linha, 532).setValue(Entrada.getRange('D145').getValue());
      SheetBanco.getRange(Linha, 533).setValue(Entrada.getRange('D146').getValue());
      SheetBanco.getRange(Linha, 534).setValue(Entrada.getRange('D147').getValue());
      SheetBanco.getRange(Linha, 535).setValue(Entrada.getRange('D148').getValue());
      SheetBanco.getRange(Linha, 536).setValue(Entrada.getRange('D149').getValue());
      SheetBanco.getRange(Linha, 537).setValue(Entrada.getRange('D150').getValue());
      SheetBanco.getRange(Linha, 538).setValue(Entrada.getRange('D151').getValue());
      SheetBanco.getRange(Linha, 539).setValue(Entrada.getRange('D152').getValue());
      SheetBanco.getRange(Linha, 540).setValue(Entrada.getRange('D153').getValue());
      SheetBanco.getRange(Linha, 541).setValue(Entrada.getRange('D154').getValue());
      SheetBanco.getRange(Linha, 542).setValue(Entrada.getRange('D155').getValue());
      SheetBanco.getRange(Linha, 543).setValue(Entrada.getRange('D156').getValue());
      SheetBanco.getRange(Linha, 544).setValue(Entrada.getRange('D157').getValue());
      SheetBanco.getRange(Linha, 545).setValue(Entrada.getRange('D158').getValue());
      SheetBanco.getRange(Linha, 546).setValue(Entrada.getRange('D159').getValue());
      SheetBanco.getRange(Linha, 547).setValue(Entrada.getRange('D160').getValue());
      SheetBanco.getRange(Linha, 548).setValue(Entrada.getRange('D161').getValue());
      SheetBanco.getRange(Linha, 549).setValue(Entrada.getRange('D162').getValue());
      SheetBanco.getRange(Linha, 550).setValue(Entrada.getRange('D163').getValue());
      SheetBanco.getRange(Linha, 551).setValue(Entrada.getRange('D164').getValue());
      SheetBanco.getRange(Linha, 552).setValue(Entrada.getRange('D165').getValue());
      SheetBanco.getRange(Linha, 553).setValue(Entrada.getRange('D166').getValue());
      SheetBanco.getRange(Linha, 554).setValue(Entrada.getRange('D167').getValue());
      SheetBanco.getRange(Linha, 555).setValue(Entrada.getRange('D168').getValue());
      SheetBanco.getRange(Linha, 556).setValue(Entrada.getRange('D169').getValue());
      SheetBanco.getRange(Linha, 557).setValue(Entrada.getRange('D170').getValue());
      SheetBanco.getRange(Linha, 558).setValue(Entrada.getRange('D171').getValue());
      SheetBanco.getRange(Linha, 559).setValue(Entrada.getRange('D172').getValue());
      SheetBanco.getRange(Linha, 560).setValue(Entrada.getRange('D173').getValue());
      SheetBanco.getRange(Linha, 561).setValue(Entrada.getRange('D174').getValue());
      SheetBanco.getRange(Linha, 562).setValue(Entrada.getRange('D175').getValue());
      SheetBanco.getRange(Linha, 563).setValue(Entrada.getRange('D176').getValue());
      SheetBanco.getRange(Linha, 564).setValue(Entrada.getRange('D177').getValue());
      SheetBanco.getRange(Linha, 565).setValue(Entrada.getRange('D178').getValue());
      SheetBanco.getRange(Linha, 566).setValue(Entrada.getRange('D179').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('D180').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('D181').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('D182').getValue());
      SheetBanco.getRange(Linha, 567).setValue(Entrada.getRange('F137').getValue());
      SheetBanco.getRange(Linha, 568).setValue(Entrada.getRange('F138').getValue());
      SheetBanco.getRange(Linha, 569).setValue(Entrada.getRange('F139').getValue());
      SheetBanco.getRange(Linha, 570).setValue(Entrada.getRange('F140').getValue());
      SheetBanco.getRange(Linha, 571).setValue(Entrada.getRange('F141').getValue());
      SheetBanco.getRange(Linha, 572).setValue(Entrada.getRange('F142').getValue());
      SheetBanco.getRange(Linha, 573).setValue(Entrada.getRange('F143').getValue());
      SheetBanco.getRange(Linha, 574).setValue(Entrada.getRange('F144').getValue());
      SheetBanco.getRange(Linha, 575).setValue(Entrada.getRange('F145').getValue());
      SheetBanco.getRange(Linha, 576).setValue(Entrada.getRange('F146').getValue());
      SheetBanco.getRange(Linha, 577).setValue(Entrada.getRange('F147').getValue());
      SheetBanco.getRange(Linha, 578).setValue(Entrada.getRange('F148').getValue());
      SheetBanco.getRange(Linha, 579).setValue(Entrada.getRange('F149').getValue());
      SheetBanco.getRange(Linha, 580).setValue(Entrada.getRange('F150').getValue());
      SheetBanco.getRange(Linha, 581).setValue(Entrada.getRange('F151').getValue());
      SheetBanco.getRange(Linha, 582).setValue(Entrada.getRange('F152').getValue());
      SheetBanco.getRange(Linha, 583).setValue(Entrada.getRange('F153').getValue());
      SheetBanco.getRange(Linha, 584).setValue(Entrada.getRange('F154').getValue());
      SheetBanco.getRange(Linha, 585).setValue(Entrada.getRange('F155').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('F156').getValue());
      SheetBanco.getRange(Linha, 586).setValue(Entrada.getRange('F157').getValue());
      SheetBanco.getRange(Linha, 587).setValue(Entrada.getRange('F158').getValue());
      SheetBanco.getRange(Linha, 588).setValue(Entrada.getRange('F159').getValue());
      SheetBanco.getRange(Linha, 589).setValue(Entrada.getRange('F160').getValue());
      SheetBanco.getRange(Linha, 590).setValue(Entrada.getRange('F161').getValue());
      SheetBanco.getRange(Linha, 591).setValue(Entrada.getRange('F162').getValue());
      SheetBanco.getRange(Linha, 592).setValue(Entrada.getRange('F163').getValue());
      SheetBanco.getRange(Linha, 593).setValue(Entrada.getRange('F164').getValue());
      SheetBanco.getRange(Linha, 594).setValue(Entrada.getRange('F165').getValue());
      SheetBanco.getRange(Linha, 595).setValue(Entrada.getRange('F166').getValue());
      SheetBanco.getRange(Linha, 596).setValue(Entrada.getRange('F167').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('F168').getValue());
      SheetBanco.getRange(Linha, 597).setValue(Entrada.getRange('F169').getValue());
      SheetBanco.getRange(Linha, 598).setValue(Entrada.getRange('F170').getValue());
      SheetBanco.getRange(Linha, 599).setValue(Entrada.getRange('F171').getValue());
      SheetBanco.getRange(Linha, 600).setValue(Entrada.getRange('F172').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('F173').getValue());
      SheetBanco.getRange(Linha, 601).setValue(Entrada.getRange('F174').getValue());
      SheetBanco.getRange(Linha, 602).setValue(Entrada.getRange('F175').getValue());
      SheetBanco.getRange(Linha, 603).setValue(Entrada.getRange('F176').getValue());
      SheetBanco.getRange(Linha, 604).setValue(Entrada.getRange('F177').getValue());
      SheetBanco.getRange(Linha, 605).setValue(Entrada.getRange('F178').getValue());
      SheetBanco.getRange(Linha, 606).setValue(Entrada.getRange('F179').getValue());
      SheetBanco.getRange(Linha, 607).setValue(Entrada.getRange('F180').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('F181').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('F182').getValue());
      SheetBanco.getRange(Linha, 608).setValue(Entrada.getRange('H137').getValue());
      SheetBanco.getRange(Linha, 609).setValue(Entrada.getRange('H138').getValue());
      SheetBanco.getRange(Linha, 610).setValue(Entrada.getRange('H139').getValue());
      SheetBanco.getRange(Linha, 611).setValue(Entrada.getRange('H140').getValue());
      SheetBanco.getRange(Linha, 612).setValue(Entrada.getRange('H141').getValue());
      SheetBanco.getRange(Linha, 613).setValue(Entrada.getRange('H142').getValue());
      SheetBanco.getRange(Linha, 614).setValue(Entrada.getRange('H143').getValue());
      SheetBanco.getRange(Linha, 615).setValue(Entrada.getRange('H144').getValue());
      SheetBanco.getRange(Linha, 616).setValue(Entrada.getRange('H145').getValue());
      SheetBanco.getRange(Linha, 617).setValue(Entrada.getRange('H146').getValue());
      SheetBanco.getRange(Linha, 618).setValue(Entrada.getRange('H147').getValue());
      SheetBanco.getRange(Linha, 619).setValue(Entrada.getRange('H148').getValue());
      SheetBanco.getRange(Linha, 620).setValue(Entrada.getRange('H149').getValue());
      SheetBanco.getRange(Linha, 621).setValue(Entrada.getRange('H150').getValue());
      SheetBanco.getRange(Linha, 622).setValue(Entrada.getRange('H151').getValue());
      SheetBanco.getRange(Linha, 623).setValue(Entrada.getRange('H152').getValue());
      SheetBanco.getRange(Linha, 624).setValue(Entrada.getRange('H153').getValue());
      SheetBanco.getRange(Linha, 625).setValue(Entrada.getRange('H154').getValue());
      SheetBanco.getRange(Linha, 626).setValue(Entrada.getRange('H155').getValue());
      SheetBanco.getRange(Linha, 627).setValue(Entrada.getRange('H156').getValue());
      SheetBanco.getRange(Linha, 628).setValue(Entrada.getRange('H157').getValue());
      SheetBanco.getRange(Linha, 629).setValue(Entrada.getRange('H158').getValue());
      SheetBanco.getRange(Linha, 630).setValue(Entrada.getRange('H159').getValue());
      SheetBanco.getRange(Linha, 631).setValue(Entrada.getRange('H160').getValue());
      SheetBanco.getRange(Linha, 632).setValue(Entrada.getRange('H161').getValue());
      SheetBanco.getRange(Linha, 633).setValue(Entrada.getRange('H162').getValue());
      SheetBanco.getRange(Linha, 634).setValue(Entrada.getRange('H163').getValue());
      SheetBanco.getRange(Linha, 635).setValue(Entrada.getRange('H164').getValue());
      SheetBanco.getRange(Linha, 636).setValue(Entrada.getRange('H165').getValue());
      SheetBanco.getRange(Linha, 637).setValue(Entrada.getRange('H166').getValue());
      SheetBanco.getRange(Linha, 638).setValue(Entrada.getRange('H167').getValue());
      SheetBanco.getRange(Linha, 639).setValue(Entrada.getRange('H168').getValue());
      SheetBanco.getRange(Linha, 640).setValue(Entrada.getRange('H169').getValue());
      SheetBanco.getRange(Linha, 641).setValue(Entrada.getRange('H170').getValue());
      SheetBanco.getRange(Linha, 642).setValue(Entrada.getRange('H171').getValue());
      SheetBanco.getRange(Linha, 643).setValue(Entrada.getRange('H172').getValue());
      SheetBanco.getRange(Linha, 644).setValue(Entrada.getRange('H173').getValue());
      SheetBanco.getRange(Linha, 645).setValue(Entrada.getRange('H174').getValue());
      SheetBanco.getRange(Linha, 646).setValue(Entrada.getRange('H175').getValue());
      SheetBanco.getRange(Linha, 647).setValue(Entrada.getRange('H176').getValue());
      SheetBanco.getRange(Linha, 648).setValue(Entrada.getRange('H177').getValue());
      SheetBanco.getRange(Linha, 649).setValue(Entrada.getRange('H178').getValue());
      SheetBanco.getRange(Linha, 650).setValue(Entrada.getRange('H179').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('H180').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('H181').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('H182').getValue());
      
       Browser.msgBox('Ocorrência Editada!')
      
      Entrada.getRangeList(['B2:B16', 'D2:D6', 'D8', 'D11','D13', 'F8:F17', 'H2:H13', 'H14', 'H16', 'B19:B23', 'B27:B37', 'B39:B41', 'D19:D26', 'D28:D31', 'D33:D39', 'F19:F41', 'H19:H38', 'B43:B47', 'B51:B61', 'B63:B73', 'B75:B78', 'B80:B86', 'D43:D85', 'F43:F47', 'F51:F61', 'F63:F73', 'F75:F78', 'F80:F86', 'H43:H85', 'B90:B94', 'B98:B108', 'B110:B120', 'B122:B125', 'B127:B133', 'D90:D132', 'F90:F94', 'F98:F108', 'F110:F120', 'F122:F125', 'F127:F133', 'H90:H132', 'B137:B141', 'B145:B155', 'B157:B167', 'B169:B172', 'B174:B180', 'D137:D179', 'F137:F141', 'F145:F155', 'F157:F167', 'F169:F172', 'F174:F180', 'H137:H179']).activate();
      Entrada.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  
      Entrada.getRange('B2').activate();
    
       }else{
    Browser.msgBox("Preencha todos os campos!")
                    }
  } else {
    
    Browser.msgBox('Ocorrência não localizada!')
    
  }

}

                 
  
  
  function LocalizarCelulaVazia(){
    
    var SheetBanco = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Banco');
    //var Entrada = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Entrada');
    
    SheetBanco.getActiveCell();
    SheetBanco.getRange('A1').activate();
    
    SheetBanco.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate();
    
    var Linha = SheetBanco.getCurrentCell().getRow();
    var NL = SheetBanco.getRange(Linha, 1).getValue();
    
    NL = NL + 1
    
    SheetBanco.getActiveCell().offset(1,0).activate()
    
    var LinhaR = Linha + 1
    
    //Entrada.getRange(2,2).setValue(NL);
    SheetBanco.getRange(LinhaR, 1).setValue(NL);
    
  }
  
  function LocalizarCelulaVazia_Excluida(){
    
    var Excluidas = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Excluidas');
     
    Excluidas.getActiveCell();
    Excluidas.getRange('A1').activate();
    
    Excluidas.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate();
    
    var Linha = Excluidas.getCurrentCell().getRow();
    var NL2 = Excluidas.getRange(Linha, 1).getValue();
    
    NL2 = NL2 + 1
    
    Excluidas.getActiveCell().offset(1,0).activate()
    
    var LinhaE = Linha + 1
    
    //Entrada.getRange(2,2).setValue(NL);
    Excluidas.getRange(LinhaE, 1).setValue(NL2);
    
   
  }
  
  function SalvarDados(){
    
    var Entrada = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Entrada'); 
    
    if(Entrada.getRange(2, 2).getValue() == ""){
      Browser.msgBox("Preencher campo Nº CONTROLE");
      Entrada.getRange('B2').activate();
      return false
    }
    
    var verif = Verificar();
    
    if(verif == true){
      LocalizarCelulaVazia();
    }else{
      Browser.msgBox("Preencha todos os campos!")
    }
    
    var SheetBanco = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Banco');
    var Linha = SheetBanco.getCurrentCell().getRow();
    
    var LocalPesquisa = SheetBanco.getRange(2, 2, SheetBanco.getLastRow()).getValues();
    var Resultado = LocalPesquisa.Pesquisa(Entrada.getRange('B2').getValue());
    
    if(Resultado !== -1){
    
      Browser.msgBox("Essa ocorrencia já está cadastrada!")
      Entrada.getRange('B2').activate();
      
      return false
    }
    
      SheetBanco.getRange(Linha, 2).setValue(Entrada.getRange('B2').getValue());
      SheetBanco.getRange(Linha, 3).setValue(Entrada.getRange('B3').getValue());
      SheetBanco.getRange(Linha, 4).setValue(Entrada.getRange('B4').getValue());
      SheetBanco.getRange(Linha, 5).setValue(Entrada.getRange('B5').getValue());
      SheetBanco.getRange(Linha, 6).setValue(Entrada.getRange('B6').getValue());
      SheetBanco.getRange(Linha, 7).setValue(Entrada.getRange('B7').getValue());
      SheetBanco.getRange(Linha, 8).setValue(Entrada.getRange('B8').getValue());
      SheetBanco.getRange(Linha, 9).setValue(Entrada.getRange('B9').getValue());
      SheetBanco.getRange(Linha, 10).setValue(Entrada.getRange('B10').getValue());
      SheetBanco.getRange(Linha, 11).setValue(Entrada.getRange('B11').getValue());
      SheetBanco.getRange(Linha, 12).setValue(Entrada.getRange('B12').getValue());
      SheetBanco.getRange(Linha, 13).setValue(Entrada.getRange('B13').getValue());
      SheetBanco.getRange(Linha, 14).setValue(Entrada.getRange('B14').getValue());
      SheetBanco.getRange(Linha, 15).setValue(Entrada.getRange('B15').getValue());
      SheetBanco.getRange(Linha, 16).setValue(Entrada.getRange('B16').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('B17').getValue());
      SheetBanco.getRange(Linha, 17).setValue(Entrada.getRange('D2').getValue());
      SheetBanco.getRange(Linha, 18).setValue(Entrada.getRange('D3').getValue());
      SheetBanco.getRange(Linha, 19).setValue(Entrada.getRange('D4').getValue());
      SheetBanco.getRange(Linha, 20).setValue(Entrada.getRange('D5').getValue());
      SheetBanco.getRange(Linha, 21).setValue(Entrada.getRange('D6').getValue());
      SheetBanco.getRange(Linha, 22).setValue(Entrada.getRange('D7').getValue());
      SheetBanco.getRange(Linha, 23).setValue(Entrada.getRange('D8').getValue());
      SheetBanco.getRange(Linha, 24).setValue(Entrada.getRange('D9').getValue());
      SheetBanco.getRange(Linha, 25).setValue(Entrada.getRange('D10').getValue());
      SheetBanco.getRange(Linha, 26).setValue(Entrada.getRange('D11').getValue());
      SheetBanco.getRange(Linha, 27).setValue(Entrada.getRange('D12').getValue());
      SheetBanco.getRange(Linha, 28).setValue(Entrada.getRange('D13').getValue());
      SheetBanco.getRange(Linha, 29).setValue(Entrada.getRange('D14').getValue());
      SheetBanco.getRange(Linha, 30).setValue(Entrada.getRange('D15').getValue());
      SheetBanco.getRange(Linha, 31).setValue(Entrada.getRange('D16').getValue());
      SheetBanco.getRange(Linha, 32).setValue(Entrada.getRange('D17').getValue());
      SheetBanco.getRange(Linha, 33).setValue(Entrada.getRange('F2').getValue());
      SheetBanco.getRange(Linha, 34).setValue(Entrada.getRange('F3').getValue());
      SheetBanco.getRange(Linha, 35).setValue(Entrada.getRange('F4').getValue());
      SheetBanco.getRange(Linha, 36).setValue(Entrada.getRange('F5').getValue());
      SheetBanco.getRange(Linha, 37).setValue(Entrada.getRange('F6').getValue());
      SheetBanco.getRange(Linha, 38).setValue(Entrada.getRange('F7').getValue());
      SheetBanco.getRange(Linha, 39).setValue(Entrada.getRange('F8').getValue());
      SheetBanco.getRange(Linha, 40).setValue(Entrada.getRange('F9').getValue());
      SheetBanco.getRange(Linha, 41).setValue(Entrada.getRange('F10').getValue());
      SheetBanco.getRange(Linha, 42).setValue(Entrada.getRange('F11').getValue());
      SheetBanco.getRange(Linha, 43).setValue(Entrada.getRange('F12').getValue());
      SheetBanco.getRange(Linha, 44).setValue(Entrada.getRange('F13').getValue());
      SheetBanco.getRange(Linha, 45).setValue(Entrada.getRange('F14').getValue());
      SheetBanco.getRange(Linha, 46).setValue(Entrada.getRange('F15').getValue());
      SheetBanco.getRange(Linha, 47).setValue(Entrada.getRange('F16').getValue());
      SheetBanco.getRange(Linha, 48).setValue(Entrada.getRange('F17').getValue());
      SheetBanco.getRange(Linha, 49).setValue(Entrada.getRange('H2').getValue());
      SheetBanco.getRange(Linha, 50).setValue(Entrada.getRange('H3').getValue());
      SheetBanco.getRange(Linha, 51).setValue(Entrada.getRange('H4').getValue());
      SheetBanco.getRange(Linha, 52).setValue(Entrada.getRange('H5').getValue());
      SheetBanco.getRange(Linha, 53).setValue(Entrada.getRange('H6').getValue());
      SheetBanco.getRange(Linha, 54).setValue(Entrada.getRange('H7').getValue());
      SheetBanco.getRange(Linha, 55).setValue(Entrada.getRange('H8').getValue());
      SheetBanco.getRange(Linha, 56).setValue(Entrada.getRange('H9').getValue());
      SheetBanco.getRange(Linha, 57).setValue(Entrada.getRange('H10').getValue());
      SheetBanco.getRange(Linha, 58).setValue(Entrada.getRange('H11').getValue());
      SheetBanco.getRange(Linha, 59).setValue(Entrada.getRange('H12').getValue());
      SheetBanco.getRange(Linha, 60).setValue(Entrada.getRange('H13').getValue());
      SheetBanco.getRange(Linha, 61).setValue(Entrada.getRange('H14').getValue());
      SheetBanco.getRange(Linha, 62).setValue(Entrada.getRange('H16').getValue());
      SheetBanco.getRange(Linha, 63).setValue(Entrada.getRange('B19').getValue());
      SheetBanco.getRange(Linha, 64).setValue(Entrada.getRange('B20').getValue());
      SheetBanco.getRange(Linha, 65).setValue(Entrada.getRange('B21').getValue());
      SheetBanco.getRange(Linha, 66).setValue(Entrada.getRange('B22').getValue());
      SheetBanco.getRange(Linha, 67).setValue(Entrada.getRange('B23').getValue());
      SheetBanco.getRange(Linha, 68).setValue(Entrada.getRange('B24').getValue());
      SheetBanco.getRange(Linha, 69).setValue(Entrada.getRange('B25').getValue());
      SheetBanco.getRange(Linha, 70).setValue(Entrada.getRange('B26').getValue());
      SheetBanco.getRange(Linha, 71).setValue(Entrada.getRange('B27').getValue());
      SheetBanco.getRange(Linha, 72).setValue(Entrada.getRange('B28').getValue());
      SheetBanco.getRange(Linha, 73).setValue(Entrada.getRange('B29').getValue());
      SheetBanco.getRange(Linha, 74).setValue(Entrada.getRange('B30').getValue());
      SheetBanco.getRange(Linha, 75).setValue(Entrada.getRange('B31').getValue());
      SheetBanco.getRange(Linha, 76).setValue(Entrada.getRange('B32').getValue());
      SheetBanco.getRange(Linha, 77).setValue(Entrada.getRange('B33').getValue());
      SheetBanco.getRange(Linha, 78).setValue(Entrada.getRange('B34').getValue());
      SheetBanco.getRange(Linha, 79).setValue(Entrada.getRange('B35').getValue());
      SheetBanco.getRange(Linha, 80).setValue(Entrada.getRange('B36').getValue());
      SheetBanco.getRange(Linha, 81).setValue(Entrada.getRange('B37').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('B38').getValue());
      SheetBanco.getRange(Linha, 82).setValue(Entrada.getRange('B39').getValue());
      SheetBanco.getRange(Linha, 83).setValue(Entrada.getRange('B40').getValue());
      SheetBanco.getRange(Linha, 84).setValue(Entrada.getRange('B41').getValue());
      SheetBanco.getRange(Linha, 85).setValue(Entrada.getRange('D19').getValue());
      SheetBanco.getRange(Linha, 86).setValue(Entrada.getRange('D20').getValue());
      SheetBanco.getRange(Linha, 87).setValue(Entrada.getRange('D21').getValue());
      SheetBanco.getRange(Linha, 88).setValue(Entrada.getRange('D22').getValue());
      SheetBanco.getRange(Linha, 89).setValue(Entrada.getRange('D23').getValue());
      SheetBanco.getRange(Linha, 90).setValue(Entrada.getRange('D24').getValue());
      SheetBanco.getRange(Linha, 91).setValue(Entrada.getRange('D25').getValue());
      SheetBanco.getRange(Linha, 92).setValue(Entrada.getRange('D26').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('D27').getValue());
      SheetBanco.getRange(Linha, 93).setValue(Entrada.getRange('D28').getValue());
      SheetBanco.getRange(Linha, 94).setValue(Entrada.getRange('D29').getValue());
      SheetBanco.getRange(Linha, 95).setValue(Entrada.getRange('D30').getValue());
      SheetBanco.getRange(Linha, 96).setValue(Entrada.getRange('D31').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('D32').getValue());
      SheetBanco.getRange(Linha, 97).setValue(Entrada.getRange('D33').getValue());
      SheetBanco.getRange(Linha, 98).setValue(Entrada.getRange('D34').getValue());
      SheetBanco.getRange(Linha, 99).setValue(Entrada.getRange('D35').getValue());
      SheetBanco.getRange(Linha, 100).setValue(Entrada.getRange('D36').getValue());
      SheetBanco.getRange(Linha, 101).setValue(Entrada.getRange('D37').getValue());
      SheetBanco.getRange(Linha, 102).setValue(Entrada.getRange('D38').getValue());
      SheetBanco.getRange(Linha, 103).setValue(Entrada.getRange('D39').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('D40').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('D41').getValue());
      SheetBanco.getRange(Linha, 104).setValue(Entrada.getRange('F19').getValue());
      SheetBanco.getRange(Linha, 105).setValue(Entrada.getRange('F20').getValue());
      SheetBanco.getRange(Linha, 106).setValue(Entrada.getRange('F21').getValue());
      SheetBanco.getRange(Linha, 107).setValue(Entrada.getRange('F22').getValue());
      SheetBanco.getRange(Linha, 108).setValue(Entrada.getRange('F23').getValue());
      SheetBanco.getRange(Linha, 109).setValue(Entrada.getRange('F24').getValue());
      SheetBanco.getRange(Linha, 110).setValue(Entrada.getRange('F25').getValue());
      SheetBanco.getRange(Linha, 111).setValue(Entrada.getRange('F26').getValue());
      SheetBanco.getRange(Linha, 112).setValue(Entrada.getRange('F27').getValue());
      SheetBanco.getRange(Linha, 113).setValue(Entrada.getRange('F28').getValue());
      SheetBanco.getRange(Linha, 114).setValue(Entrada.getRange('F29').getValue());
      SheetBanco.getRange(Linha, 115).setValue(Entrada.getRange('F30').getValue());
      SheetBanco.getRange(Linha, 116).setValue(Entrada.getRange('F31').getValue());
      SheetBanco.getRange(Linha, 117).setValue(Entrada.getRange('F32').getValue());
      SheetBanco.getRange(Linha, 118).setValue(Entrada.getRange('F33').getValue());
      SheetBanco.getRange(Linha, 119).setValue(Entrada.getRange('F34').getValue());
      SheetBanco.getRange(Linha, 120).setValue(Entrada.getRange('F35').getValue());
      SheetBanco.getRange(Linha, 121).setValue(Entrada.getRange('F36').getValue());
      SheetBanco.getRange(Linha, 122).setValue(Entrada.getRange('F37').getValue());
      SheetBanco.getRange(Linha, 123).setValue(Entrada.getRange('F38').getValue());
      SheetBanco.getRange(Linha, 124).setValue(Entrada.getRange('F39').getValue());
      SheetBanco.getRange(Linha, 125).setValue(Entrada.getRange('F40').getValue());
      SheetBanco.getRange(Linha, 126).setValue(Entrada.getRange('F41').getValue());
      SheetBanco.getRange(Linha, 127).setValue(Entrada.getRange('H19').getValue());
      SheetBanco.getRange(Linha, 128).setValue(Entrada.getRange('H20').getValue());
      SheetBanco.getRange(Linha, 129).setValue(Entrada.getRange('H21').getValue());
      SheetBanco.getRange(Linha, 130).setValue(Entrada.getRange('H22').getValue());
      SheetBanco.getRange(Linha, 131).setValue(Entrada.getRange('H23').getValue());
      SheetBanco.getRange(Linha, 132).setValue(Entrada.getRange('H24').getValue());
      SheetBanco.getRange(Linha, 133).setValue(Entrada.getRange('H25').getValue());
      SheetBanco.getRange(Linha, 134).setValue(Entrada.getRange('H26').getValue());
      SheetBanco.getRange(Linha, 135).setValue(Entrada.getRange('H27').getValue());
      SheetBanco.getRange(Linha, 136).setValue(Entrada.getRange('H28').getValue());
      SheetBanco.getRange(Linha, 137).setValue(Entrada.getRange('H29').getValue());
      SheetBanco.getRange(Linha, 138).setValue(Entrada.getRange('H30').getValue());
      SheetBanco.getRange(Linha, 139).setValue(Entrada.getRange('H31').getValue());
      SheetBanco.getRange(Linha, 140).setValue(Entrada.getRange('H32').getValue());
      SheetBanco.getRange(Linha, 141).setValue(Entrada.getRange('H33').getValue());
      SheetBanco.getRange(Linha, 142).setValue(Entrada.getRange('H34').getValue());
      SheetBanco.getRange(Linha, 143).setValue(Entrada.getRange('H35').getValue());
      SheetBanco.getRange(Linha, 144).setValue(Entrada.getRange('H36').getValue());
      SheetBanco.getRange(Linha, 145).setValue(Entrada.getRange('H37').getValue());
      SheetBanco.getRange(Linha, 146).setValue(Entrada.getRange('H38').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('H39').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('H40').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('H41').getValue());
      SheetBanco.getRange(Linha, 147).setValue(Entrada.getRange('B43').getValue());
      SheetBanco.getRange(Linha, 148).setValue(Entrada.getRange('B44').getValue());
      SheetBanco.getRange(Linha, 149).setValue(Entrada.getRange('B45').getValue());
      SheetBanco.getRange(Linha, 150).setValue(Entrada.getRange('B46').getValue());
      SheetBanco.getRange(Linha, 151).setValue(Entrada.getRange('B47').getValue());
      SheetBanco.getRange(Linha, 152).setValue(Entrada.getRange('B48').getValue());
      SheetBanco.getRange(Linha, 153).setValue(Entrada.getRange('B49').getValue());
      SheetBanco.getRange(Linha, 154).setValue(Entrada.getRange('B50').getValue());
      SheetBanco.getRange(Linha, 155).setValue(Entrada.getRange('B51').getValue());
      SheetBanco.getRange(Linha, 156).setValue(Entrada.getRange('B52').getValue());
      SheetBanco.getRange(Linha, 157).setValue(Entrada.getRange('B53').getValue());
      SheetBanco.getRange(Linha, 158).setValue(Entrada.getRange('B54').getValue());
      SheetBanco.getRange(Linha, 159).setValue(Entrada.getRange('B55').getValue());
      SheetBanco.getRange(Linha, 160).setValue(Entrada.getRange('B56').getValue());
      SheetBanco.getRange(Linha, 161).setValue(Entrada.getRange('B57').getValue());
      SheetBanco.getRange(Linha, 162).setValue(Entrada.getRange('B58').getValue());
      SheetBanco.getRange(Linha, 163).setValue(Entrada.getRange('B59').getValue());
      SheetBanco.getRange(Linha, 164).setValue(Entrada.getRange('B60').getValue());
      SheetBanco.getRange(Linha, 165).setValue(Entrada.getRange('B61').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('B62').getValue());
      SheetBanco.getRange(Linha, 166).setValue(Entrada.getRange('B63').getValue());
      SheetBanco.getRange(Linha, 167).setValue(Entrada.getRange('B64').getValue());
      SheetBanco.getRange(Linha, 168).setValue(Entrada.getRange('B65').getValue());
      SheetBanco.getRange(Linha, 169).setValue(Entrada.getRange('B66').getValue());
      SheetBanco.getRange(Linha, 170).setValue(Entrada.getRange('B67').getValue());
      SheetBanco.getRange(Linha, 171).setValue(Entrada.getRange('B68').getValue());
      SheetBanco.getRange(Linha, 172).setValue(Entrada.getRange('B69').getValue());
      SheetBanco.getRange(Linha, 173).setValue(Entrada.getRange('B70').getValue());
      SheetBanco.getRange(Linha, 174).setValue(Entrada.getRange('B71').getValue());
      SheetBanco.getRange(Linha, 175).setValue(Entrada.getRange('B72').getValue());
      SheetBanco.getRange(Linha, 176).setValue(Entrada.getRange('B73').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('B74').getValue());
      SheetBanco.getRange(Linha, 177).setValue(Entrada.getRange('B75').getValue());
      SheetBanco.getRange(Linha, 178).setValue(Entrada.getRange('B76').getValue());
      SheetBanco.getRange(Linha, 179).setValue(Entrada.getRange('B77').getValue());
      SheetBanco.getRange(Linha, 180).setValue(Entrada.getRange('B78').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('B79').getValue());
      SheetBanco.getRange(Linha, 181).setValue(Entrada.getRange('B80').getValue());
      SheetBanco.getRange(Linha, 182).setValue(Entrada.getRange('B81').getValue());
      SheetBanco.getRange(Linha, 183).setValue(Entrada.getRange('B82').getValue());
      SheetBanco.getRange(Linha, 184).setValue(Entrada.getRange('B83').getValue());
      SheetBanco.getRange(Linha, 185).setValue(Entrada.getRange('B84').getValue());
      SheetBanco.getRange(Linha, 186).setValue(Entrada.getRange('B85').getValue());
      SheetBanco.getRange(Linha, 187).setValue(Entrada.getRange('B86').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('B87').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('B88').getValue());
      SheetBanco.getRange(Linha, 188).setValue(Entrada.getRange('D43').getValue());
      SheetBanco.getRange(Linha, 189).setValue(Entrada.getRange('D44').getValue());
      SheetBanco.getRange(Linha, 190).setValue(Entrada.getRange('D45').getValue());
      SheetBanco.getRange(Linha, 191).setValue(Entrada.getRange('D46').getValue());
      SheetBanco.getRange(Linha, 192).setValue(Entrada.getRange('D47').getValue());
      SheetBanco.getRange(Linha, 193).setValue(Entrada.getRange('D48').getValue());
      SheetBanco.getRange(Linha, 194).setValue(Entrada.getRange('D49').getValue());
      SheetBanco.getRange(Linha, 195).setValue(Entrada.getRange('D50').getValue());
      SheetBanco.getRange(Linha, 196).setValue(Entrada.getRange('D51').getValue());
      SheetBanco.getRange(Linha, 197).setValue(Entrada.getRange('D52').getValue());
      SheetBanco.getRange(Linha, 198).setValue(Entrada.getRange('D53').getValue());
      SheetBanco.getRange(Linha, 199).setValue(Entrada.getRange('D54').getValue());
      SheetBanco.getRange(Linha, 200).setValue(Entrada.getRange('D55').getValue());
      SheetBanco.getRange(Linha, 201).setValue(Entrada.getRange('D56').getValue());
      SheetBanco.getRange(Linha, 202).setValue(Entrada.getRange('D57').getValue());
      SheetBanco.getRange(Linha, 203).setValue(Entrada.getRange('D58').getValue());
      SheetBanco.getRange(Linha, 204).setValue(Entrada.getRange('D59').getValue());
      SheetBanco.getRange(Linha, 205).setValue(Entrada.getRange('D60').getValue());
      SheetBanco.getRange(Linha, 206).setValue(Entrada.getRange('D61').getValue());
      SheetBanco.getRange(Linha, 207).setValue(Entrada.getRange('D62').getValue());
      SheetBanco.getRange(Linha, 208).setValue(Entrada.getRange('D63').getValue());
      SheetBanco.getRange(Linha, 209).setValue(Entrada.getRange('D64').getValue());
      SheetBanco.getRange(Linha, 210).setValue(Entrada.getRange('D65').getValue());
      SheetBanco.getRange(Linha, 211).setValue(Entrada.getRange('D66').getValue());
      SheetBanco.getRange(Linha, 212).setValue(Entrada.getRange('D67').getValue());
      SheetBanco.getRange(Linha, 213).setValue(Entrada.getRange('D68').getValue());
      SheetBanco.getRange(Linha, 214).setValue(Entrada.getRange('D69').getValue());
      SheetBanco.getRange(Linha, 215).setValue(Entrada.getRange('D70').getValue());
      SheetBanco.getRange(Linha, 216).setValue(Entrada.getRange('D71').getValue());
      SheetBanco.getRange(Linha, 217).setValue(Entrada.getRange('D72').getValue());
      SheetBanco.getRange(Linha, 218).setValue(Entrada.getRange('D73').getValue());
      SheetBanco.getRange(Linha, 219).setValue(Entrada.getRange('D74').getValue());
      SheetBanco.getRange(Linha, 220).setValue(Entrada.getRange('D75').getValue());
      SheetBanco.getRange(Linha, 221).setValue(Entrada.getRange('D76').getValue());
      SheetBanco.getRange(Linha, 222).setValue(Entrada.getRange('D77').getValue());
      SheetBanco.getRange(Linha, 223).setValue(Entrada.getRange('D78').getValue());
      SheetBanco.getRange(Linha, 224).setValue(Entrada.getRange('D79').getValue());
      SheetBanco.getRange(Linha, 225).setValue(Entrada.getRange('D80').getValue());
      SheetBanco.getRange(Linha, 226).setValue(Entrada.getRange('D81').getValue());
      SheetBanco.getRange(Linha, 227).setValue(Entrada.getRange('D82').getValue());
      SheetBanco.getRange(Linha, 228).setValue(Entrada.getRange('D83').getValue());
      SheetBanco.getRange(Linha, 229).setValue(Entrada.getRange('D84').getValue());
      SheetBanco.getRange(Linha, 230).setValue(Entrada.getRange('D85').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('D86').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('D87').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('D88').getValue());
      SheetBanco.getRange(Linha, 231).setValue(Entrada.getRange('F43').getValue());
      SheetBanco.getRange(Linha, 232).setValue(Entrada.getRange('F44').getValue());
      SheetBanco.getRange(Linha, 233).setValue(Entrada.getRange('F45').getValue());
      SheetBanco.getRange(Linha, 234).setValue(Entrada.getRange('F46').getValue());
      SheetBanco.getRange(Linha, 235).setValue(Entrada.getRange('F47').getValue());
      SheetBanco.getRange(Linha, 236).setValue(Entrada.getRange('F48').getValue());
      SheetBanco.getRange(Linha, 237).setValue(Entrada.getRange('F49').getValue());
      SheetBanco.getRange(Linha, 238).setValue(Entrada.getRange('F50').getValue());
      SheetBanco.getRange(Linha, 239).setValue(Entrada.getRange('F51').getValue());
      SheetBanco.getRange(Linha, 240).setValue(Entrada.getRange('F52').getValue());
      SheetBanco.getRange(Linha, 241).setValue(Entrada.getRange('F53').getValue());
      SheetBanco.getRange(Linha, 242).setValue(Entrada.getRange('F54').getValue());
      SheetBanco.getRange(Linha, 243).setValue(Entrada.getRange('F55').getValue());
      SheetBanco.getRange(Linha, 244).setValue(Entrada.getRange('F56').getValue());
      SheetBanco.getRange(Linha, 245).setValue(Entrada.getRange('F57').getValue());
      SheetBanco.getRange(Linha, 246).setValue(Entrada.getRange('F58').getValue());
      SheetBanco.getRange(Linha, 247).setValue(Entrada.getRange('F59').getValue());
      SheetBanco.getRange(Linha, 248).setValue(Entrada.getRange('F60').getValue());
      SheetBanco.getRange(Linha, 249).setValue(Entrada.getRange('F61').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('F62').getValue());
      SheetBanco.getRange(Linha, 250).setValue(Entrada.getRange('F63').getValue());
      SheetBanco.getRange(Linha, 251).setValue(Entrada.getRange('F64').getValue());
      SheetBanco.getRange(Linha, 252).setValue(Entrada.getRange('F65').getValue());
      SheetBanco.getRange(Linha, 253).setValue(Entrada.getRange('F66').getValue());
      SheetBanco.getRange(Linha, 254).setValue(Entrada.getRange('F67').getValue());
      SheetBanco.getRange(Linha, 255).setValue(Entrada.getRange('F68').getValue());
      SheetBanco.getRange(Linha, 256).setValue(Entrada.getRange('F69').getValue());
      SheetBanco.getRange(Linha, 257).setValue(Entrada.getRange('F70').getValue());
      SheetBanco.getRange(Linha, 258).setValue(Entrada.getRange('F71').getValue());
      SheetBanco.getRange(Linha, 259).setValue(Entrada.getRange('F72').getValue());
      SheetBanco.getRange(Linha, 260).setValue(Entrada.getRange('F73').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('F74').getValue());
      SheetBanco.getRange(Linha, 261).setValue(Entrada.getRange('F75').getValue());
      SheetBanco.getRange(Linha, 262).setValue(Entrada.getRange('F76').getValue());
      SheetBanco.getRange(Linha, 263).setValue(Entrada.getRange('F77').getValue());
      SheetBanco.getRange(Linha, 264).setValue(Entrada.getRange('F78').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('F79').getValue());
      SheetBanco.getRange(Linha, 265).setValue(Entrada.getRange('F80').getValue());
      SheetBanco.getRange(Linha, 266).setValue(Entrada.getRange('F81').getValue());
      SheetBanco.getRange(Linha, 267).setValue(Entrada.getRange('F82').getValue());
      SheetBanco.getRange(Linha, 268).setValue(Entrada.getRange('F83').getValue());
      SheetBanco.getRange(Linha, 269).setValue(Entrada.getRange('F84').getValue());
      SheetBanco.getRange(Linha, 270).setValue(Entrada.getRange('F85').getValue());
      SheetBanco.getRange(Linha, 271).setValue(Entrada.getRange('F86').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('F87').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('F88').getValue());
      SheetBanco.getRange(Linha, 272).setValue(Entrada.getRange('H43').getValue());
      SheetBanco.getRange(Linha, 273).setValue(Entrada.getRange('H44').getValue());
      SheetBanco.getRange(Linha, 274).setValue(Entrada.getRange('H45').getValue());
      SheetBanco.getRange(Linha, 275).setValue(Entrada.getRange('H46').getValue());
      SheetBanco.getRange(Linha, 276).setValue(Entrada.getRange('H47').getValue());
      SheetBanco.getRange(Linha, 277).setValue(Entrada.getRange('H48').getValue());
      SheetBanco.getRange(Linha, 278).setValue(Entrada.getRange('H49').getValue());
      SheetBanco.getRange(Linha, 279).setValue(Entrada.getRange('H50').getValue());
      SheetBanco.getRange(Linha, 280).setValue(Entrada.getRange('H51').getValue());
      SheetBanco.getRange(Linha, 281).setValue(Entrada.getRange('H52').getValue());
      SheetBanco.getRange(Linha, 282).setValue(Entrada.getRange('H53').getValue());
      SheetBanco.getRange(Linha, 283).setValue(Entrada.getRange('H54').getValue());
      SheetBanco.getRange(Linha, 284).setValue(Entrada.getRange('H55').getValue());
      SheetBanco.getRange(Linha, 285).setValue(Entrada.getRange('H56').getValue());
      SheetBanco.getRange(Linha, 286).setValue(Entrada.getRange('H57').getValue());
      SheetBanco.getRange(Linha, 287).setValue(Entrada.getRange('H58').getValue());
      SheetBanco.getRange(Linha, 288).setValue(Entrada.getRange('H59').getValue());
      SheetBanco.getRange(Linha, 289).setValue(Entrada.getRange('H60').getValue());
      SheetBanco.getRange(Linha, 290).setValue(Entrada.getRange('H61').getValue());
      SheetBanco.getRange(Linha, 291).setValue(Entrada.getRange('H62').getValue());
      SheetBanco.getRange(Linha, 292).setValue(Entrada.getRange('H63').getValue());
      SheetBanco.getRange(Linha, 293).setValue(Entrada.getRange('H64').getValue());
      SheetBanco.getRange(Linha, 294).setValue(Entrada.getRange('H65').getValue());
      SheetBanco.getRange(Linha, 295).setValue(Entrada.getRange('H66').getValue());
      SheetBanco.getRange(Linha, 296).setValue(Entrada.getRange('H67').getValue());
      SheetBanco.getRange(Linha, 297).setValue(Entrada.getRange('H68').getValue());
      SheetBanco.getRange(Linha, 298).setValue(Entrada.getRange('H69').getValue());
      SheetBanco.getRange(Linha, 299).setValue(Entrada.getRange('H70').getValue());
      SheetBanco.getRange(Linha, 300).setValue(Entrada.getRange('H71').getValue());
      SheetBanco.getRange(Linha, 301).setValue(Entrada.getRange('H72').getValue());
      SheetBanco.getRange(Linha, 302).setValue(Entrada.getRange('H73').getValue());
      SheetBanco.getRange(Linha, 303).setValue(Entrada.getRange('H74').getValue());
      SheetBanco.getRange(Linha, 304).setValue(Entrada.getRange('H75').getValue());
      SheetBanco.getRange(Linha, 305).setValue(Entrada.getRange('H76').getValue());
      SheetBanco.getRange(Linha, 306).setValue(Entrada.getRange('H77').getValue());
      SheetBanco.getRange(Linha, 307).setValue(Entrada.getRange('H78').getValue());
      SheetBanco.getRange(Linha, 308).setValue(Entrada.getRange('H79').getValue());
      SheetBanco.getRange(Linha, 309).setValue(Entrada.getRange('H80').getValue());
      SheetBanco.getRange(Linha, 310).setValue(Entrada.getRange('H81').getValue());
      SheetBanco.getRange(Linha, 311).setValue(Entrada.getRange('H82').getValue());
      SheetBanco.getRange(Linha, 312).setValue(Entrada.getRange('H83').getValue());
      SheetBanco.getRange(Linha, 313).setValue(Entrada.getRange('H84').getValue());
      SheetBanco.getRange(Linha, 314).setValue(Entrada.getRange('H85').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('H86').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('H87').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('H88').getValue());
      SheetBanco.getRange(Linha, 315).setValue(Entrada.getRange('B90').getValue());
      SheetBanco.getRange(Linha, 316).setValue(Entrada.getRange('B91').getValue());
      SheetBanco.getRange(Linha, 317).setValue(Entrada.getRange('B92').getValue());
      SheetBanco.getRange(Linha, 318).setValue(Entrada.getRange('B93').getValue());
      SheetBanco.getRange(Linha, 319).setValue(Entrada.getRange('B94').getValue());
      SheetBanco.getRange(Linha, 320).setValue(Entrada.getRange('B95').getValue());
      SheetBanco.getRange(Linha, 321).setValue(Entrada.getRange('B96').getValue());
      SheetBanco.getRange(Linha, 322).setValue(Entrada.getRange('B97').getValue());
      SheetBanco.getRange(Linha, 323).setValue(Entrada.getRange('B98').getValue());
      SheetBanco.getRange(Linha, 324).setValue(Entrada.getRange('B99').getValue());
      SheetBanco.getRange(Linha, 325).setValue(Entrada.getRange('B100').getValue());
      SheetBanco.getRange(Linha, 326).setValue(Entrada.getRange('B101').getValue());
      SheetBanco.getRange(Linha, 327).setValue(Entrada.getRange('B102').getValue());
      SheetBanco.getRange(Linha, 328).setValue(Entrada.getRange('B103').getValue());
      SheetBanco.getRange(Linha, 329).setValue(Entrada.getRange('B104').getValue());
      SheetBanco.getRange(Linha, 330).setValue(Entrada.getRange('B105').getValue());
      SheetBanco.getRange(Linha, 331).setValue(Entrada.getRange('B106').getValue());
      SheetBanco.getRange(Linha, 332).setValue(Entrada.getRange('B107').getValue());
      SheetBanco.getRange(Linha, 333).setValue(Entrada.getRange('B108').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('B109').getValue());
      SheetBanco.getRange(Linha, 334).setValue(Entrada.getRange('B110').getValue());
      SheetBanco.getRange(Linha, 335).setValue(Entrada.getRange('B111').getValue());
      SheetBanco.getRange(Linha, 336).setValue(Entrada.getRange('B112').getValue());
      SheetBanco.getRange(Linha, 337).setValue(Entrada.getRange('B113').getValue());
      SheetBanco.getRange(Linha, 338).setValue(Entrada.getRange('B114').getValue());
      SheetBanco.getRange(Linha, 339).setValue(Entrada.getRange('B115').getValue());
      SheetBanco.getRange(Linha, 340).setValue(Entrada.getRange('B116').getValue());
      SheetBanco.getRange(Linha, 341).setValue(Entrada.getRange('B117').getValue());
      SheetBanco.getRange(Linha, 342).setValue(Entrada.getRange('B118').getValue());
      SheetBanco.getRange(Linha, 343).setValue(Entrada.getRange('B119').getValue());
      SheetBanco.getRange(Linha, 344).setValue(Entrada.getRange('B120').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('B121').getValue());
      SheetBanco.getRange(Linha, 345).setValue(Entrada.getRange('B122').getValue());
      SheetBanco.getRange(Linha, 346).setValue(Entrada.getRange('B123').getValue());
      SheetBanco.getRange(Linha, 347).setValue(Entrada.getRange('B124').getValue());
      SheetBanco.getRange(Linha, 348).setValue(Entrada.getRange('B125').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('B126').getValue());
      SheetBanco.getRange(Linha, 349).setValue(Entrada.getRange('B127').getValue());
      SheetBanco.getRange(Linha, 350).setValue(Entrada.getRange('B128').getValue());
      SheetBanco.getRange(Linha, 351).setValue(Entrada.getRange('B129').getValue());
      SheetBanco.getRange(Linha, 352).setValue(Entrada.getRange('B130').getValue());
      SheetBanco.getRange(Linha, 353).setValue(Entrada.getRange('B131').getValue());
      SheetBanco.getRange(Linha, 354).setValue(Entrada.getRange('B132').getValue());
      SheetBanco.getRange(Linha, 355).setValue(Entrada.getRange('B133').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('B134').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('B135').getValue());
      SheetBanco.getRange(Linha, 356).setValue(Entrada.getRange('D90').getValue());
      SheetBanco.getRange(Linha, 357).setValue(Entrada.getRange('D91').getValue());
      SheetBanco.getRange(Linha, 358).setValue(Entrada.getRange('D92').getValue());
      SheetBanco.getRange(Linha, 359).setValue(Entrada.getRange('D93').getValue());
      SheetBanco.getRange(Linha, 360).setValue(Entrada.getRange('D94').getValue());
      SheetBanco.getRange(Linha, 361).setValue(Entrada.getRange('D95').getValue());
      SheetBanco.getRange(Linha, 362).setValue(Entrada.getRange('D96').getValue());
      SheetBanco.getRange(Linha, 363).setValue(Entrada.getRange('D97').getValue());
      SheetBanco.getRange(Linha, 364).setValue(Entrada.getRange('D98').getValue());
      SheetBanco.getRange(Linha, 365).setValue(Entrada.getRange('D99').getValue());
      SheetBanco.getRange(Linha, 366).setValue(Entrada.getRange('D100').getValue());
      SheetBanco.getRange(Linha, 367).setValue(Entrada.getRange('D101').getValue());
      SheetBanco.getRange(Linha, 368).setValue(Entrada.getRange('D102').getValue());
      SheetBanco.getRange(Linha, 369).setValue(Entrada.getRange('D103').getValue());
      SheetBanco.getRange(Linha, 370).setValue(Entrada.getRange('D104').getValue());
      SheetBanco.getRange(Linha, 371).setValue(Entrada.getRange('D105').getValue());
      SheetBanco.getRange(Linha, 372).setValue(Entrada.getRange('D106').getValue());
      SheetBanco.getRange(Linha, 373).setValue(Entrada.getRange('D107').getValue());
      SheetBanco.getRange(Linha, 374).setValue(Entrada.getRange('D108').getValue());
      SheetBanco.getRange(Linha, 375).setValue(Entrada.getRange('D109').getValue());
      SheetBanco.getRange(Linha, 376).setValue(Entrada.getRange('D110').getValue());
      SheetBanco.getRange(Linha, 377).setValue(Entrada.getRange('D111').getValue());
      SheetBanco.getRange(Linha, 378).setValue(Entrada.getRange('D112').getValue());
      SheetBanco.getRange(Linha, 379).setValue(Entrada.getRange('D113').getValue());
      SheetBanco.getRange(Linha, 380).setValue(Entrada.getRange('D114').getValue());
      SheetBanco.getRange(Linha, 381).setValue(Entrada.getRange('D115').getValue());
      SheetBanco.getRange(Linha, 382).setValue(Entrada.getRange('D116').getValue());
      SheetBanco.getRange(Linha, 383).setValue(Entrada.getRange('D117').getValue());
      SheetBanco.getRange(Linha, 384).setValue(Entrada.getRange('D118').getValue());
      SheetBanco.getRange(Linha, 385).setValue(Entrada.getRange('D119').getValue());
      SheetBanco.getRange(Linha, 386).setValue(Entrada.getRange('D120').getValue());
      SheetBanco.getRange(Linha, 387).setValue(Entrada.getRange('D121').getValue());
      SheetBanco.getRange(Linha, 388).setValue(Entrada.getRange('D122').getValue());
      SheetBanco.getRange(Linha, 389).setValue(Entrada.getRange('D123').getValue());
      SheetBanco.getRange(Linha, 390).setValue(Entrada.getRange('D124').getValue());
      SheetBanco.getRange(Linha, 391).setValue(Entrada.getRange('D125').getValue());
      SheetBanco.getRange(Linha, 392).setValue(Entrada.getRange('D126').getValue());
      SheetBanco.getRange(Linha, 393).setValue(Entrada.getRange('D127').getValue());
      SheetBanco.getRange(Linha, 394).setValue(Entrada.getRange('D128').getValue());
      SheetBanco.getRange(Linha, 395).setValue(Entrada.getRange('D129').getValue());
      SheetBanco.getRange(Linha, 396).setValue(Entrada.getRange('D130').getValue());
      SheetBanco.getRange(Linha, 397).setValue(Entrada.getRange('D131').getValue());
      SheetBanco.getRange(Linha, 398).setValue(Entrada.getRange('D132').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('D133').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('D134').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('D135').getValue());
      SheetBanco.getRange(Linha, 399).setValue(Entrada.getRange('F90').getValue());
      SheetBanco.getRange(Linha, 400).setValue(Entrada.getRange('F91').getValue());
      SheetBanco.getRange(Linha, 401).setValue(Entrada.getRange('F92').getValue());
      SheetBanco.getRange(Linha, 402).setValue(Entrada.getRange('F93').getValue());
      SheetBanco.getRange(Linha, 403).setValue(Entrada.getRange('F94').getValue());
      SheetBanco.getRange(Linha, 404).setValue(Entrada.getRange('F95').getValue());
      SheetBanco.getRange(Linha, 405).setValue(Entrada.getRange('F96').getValue());
      SheetBanco.getRange(Linha, 406).setValue(Entrada.getRange('F97').getValue());
      SheetBanco.getRange(Linha, 407).setValue(Entrada.getRange('F98').getValue());
      SheetBanco.getRange(Linha, 408).setValue(Entrada.getRange('F99').getValue());
      SheetBanco.getRange(Linha, 409).setValue(Entrada.getRange('F100').getValue());
      SheetBanco.getRange(Linha, 410).setValue(Entrada.getRange('F101').getValue());
      SheetBanco.getRange(Linha, 411).setValue(Entrada.getRange('F102').getValue());
      SheetBanco.getRange(Linha, 412).setValue(Entrada.getRange('F103').getValue());
      SheetBanco.getRange(Linha, 413).setValue(Entrada.getRange('F104').getValue());
      SheetBanco.getRange(Linha, 414).setValue(Entrada.getRange('F105').getValue());
      SheetBanco.getRange(Linha, 415).setValue(Entrada.getRange('F106').getValue());
      SheetBanco.getRange(Linha, 416).setValue(Entrada.getRange('F107').getValue());
      SheetBanco.getRange(Linha, 417).setValue(Entrada.getRange('F108').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('F109').getValue());
      SheetBanco.getRange(Linha, 418).setValue(Entrada.getRange('F110').getValue());
      SheetBanco.getRange(Linha, 419).setValue(Entrada.getRange('F111').getValue());
      SheetBanco.getRange(Linha, 420).setValue(Entrada.getRange('F112').getValue());
      SheetBanco.getRange(Linha, 421).setValue(Entrada.getRange('F113').getValue());
      SheetBanco.getRange(Linha, 422).setValue(Entrada.getRange('F114').getValue());
      SheetBanco.getRange(Linha, 423).setValue(Entrada.getRange('F115').getValue());
      SheetBanco.getRange(Linha, 424).setValue(Entrada.getRange('F116').getValue());
      SheetBanco.getRange(Linha, 425).setValue(Entrada.getRange('F117').getValue());
      SheetBanco.getRange(Linha, 426).setValue(Entrada.getRange('F118').getValue());
      SheetBanco.getRange(Linha, 427).setValue(Entrada.getRange('F119').getValue());
      SheetBanco.getRange(Linha, 428).setValue(Entrada.getRange('F120').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('F121').getValue());
      SheetBanco.getRange(Linha, 429).setValue(Entrada.getRange('F122').getValue());
      SheetBanco.getRange(Linha, 430).setValue(Entrada.getRange('F123').getValue());
      SheetBanco.getRange(Linha, 431).setValue(Entrada.getRange('F124').getValue());
      SheetBanco.getRange(Linha, 432).setValue(Entrada.getRange('F125').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('F126').getValue());
      SheetBanco.getRange(Linha, 433).setValue(Entrada.getRange('F127').getValue());
      SheetBanco.getRange(Linha, 434).setValue(Entrada.getRange('F128').getValue());
      SheetBanco.getRange(Linha, 435).setValue(Entrada.getRange('F129').getValue());
      SheetBanco.getRange(Linha, 436).setValue(Entrada.getRange('F130').getValue());
      SheetBanco.getRange(Linha, 437).setValue(Entrada.getRange('F131').getValue());
      SheetBanco.getRange(Linha, 438).setValue(Entrada.getRange('F132').getValue());
      SheetBanco.getRange(Linha, 439).setValue(Entrada.getRange('F133').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('F134').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('F135').getValue());
      SheetBanco.getRange(Linha, 440).setValue(Entrada.getRange('H90').getValue());
      SheetBanco.getRange(Linha, 441).setValue(Entrada.getRange('H91').getValue());
      SheetBanco.getRange(Linha, 442).setValue(Entrada.getRange('H92').getValue());
      SheetBanco.getRange(Linha, 443).setValue(Entrada.getRange('H93').getValue());
      SheetBanco.getRange(Linha, 444).setValue(Entrada.getRange('H94').getValue());
      SheetBanco.getRange(Linha, 445).setValue(Entrada.getRange('H95').getValue());
      SheetBanco.getRange(Linha, 446).setValue(Entrada.getRange('H96').getValue());
      SheetBanco.getRange(Linha, 447).setValue(Entrada.getRange('H97').getValue());
      SheetBanco.getRange(Linha, 448).setValue(Entrada.getRange('H98').getValue());
      SheetBanco.getRange(Linha, 449).setValue(Entrada.getRange('H99').getValue());
      SheetBanco.getRange(Linha, 450).setValue(Entrada.getRange('H100').getValue());
      SheetBanco.getRange(Linha, 451).setValue(Entrada.getRange('H101').getValue());
      SheetBanco.getRange(Linha, 452).setValue(Entrada.getRange('H102').getValue());
      SheetBanco.getRange(Linha, 453).setValue(Entrada.getRange('H103').getValue());
      SheetBanco.getRange(Linha, 454).setValue(Entrada.getRange('H104').getValue());
      SheetBanco.getRange(Linha, 455).setValue(Entrada.getRange('H105').getValue());
      SheetBanco.getRange(Linha, 456).setValue(Entrada.getRange('H106').getValue());
      SheetBanco.getRange(Linha, 457).setValue(Entrada.getRange('H107').getValue());
      SheetBanco.getRange(Linha, 458).setValue(Entrada.getRange('H108').getValue());
      SheetBanco.getRange(Linha, 459).setValue(Entrada.getRange('H109').getValue());
      SheetBanco.getRange(Linha, 460).setValue(Entrada.getRange('H110').getValue());
      SheetBanco.getRange(Linha, 461).setValue(Entrada.getRange('H111').getValue());
      SheetBanco.getRange(Linha, 462).setValue(Entrada.getRange('H112').getValue());
      SheetBanco.getRange(Linha, 463).setValue(Entrada.getRange('H113').getValue());
      SheetBanco.getRange(Linha, 464).setValue(Entrada.getRange('H114').getValue());
      SheetBanco.getRange(Linha, 465).setValue(Entrada.getRange('H115').getValue());
      SheetBanco.getRange(Linha, 466).setValue(Entrada.getRange('H116').getValue());
      SheetBanco.getRange(Linha, 467).setValue(Entrada.getRange('H117').getValue());
      SheetBanco.getRange(Linha, 468).setValue(Entrada.getRange('H118').getValue());
      SheetBanco.getRange(Linha, 469).setValue(Entrada.getRange('H119').getValue());
      SheetBanco.getRange(Linha, 470).setValue(Entrada.getRange('H120').getValue());
      SheetBanco.getRange(Linha, 471).setValue(Entrada.getRange('H121').getValue());
      SheetBanco.getRange(Linha, 472).setValue(Entrada.getRange('H122').getValue());
      SheetBanco.getRange(Linha, 473).setValue(Entrada.getRange('H123').getValue());
      SheetBanco.getRange(Linha, 474).setValue(Entrada.getRange('H124').getValue());
      SheetBanco.getRange(Linha, 475).setValue(Entrada.getRange('H125').getValue());
      SheetBanco.getRange(Linha, 476).setValue(Entrada.getRange('H126').getValue());
      SheetBanco.getRange(Linha, 477).setValue(Entrada.getRange('H127').getValue());
      SheetBanco.getRange(Linha, 478).setValue(Entrada.getRange('H128').getValue());
      SheetBanco.getRange(Linha, 479).setValue(Entrada.getRange('H129').getValue());
      SheetBanco.getRange(Linha, 480).setValue(Entrada.getRange('H130').getValue());
      SheetBanco.getRange(Linha, 481).setValue(Entrada.getRange('H131').getValue());
      SheetBanco.getRange(Linha, 482).setValue(Entrada.getRange('H132').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('H133').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('H134').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('H135').getValue());
      SheetBanco.getRange(Linha, 483).setValue(Entrada.getRange('B137').getValue());
      SheetBanco.getRange(Linha, 484).setValue(Entrada.getRange('B138').getValue());
      SheetBanco.getRange(Linha, 485).setValue(Entrada.getRange('B139').getValue());
      SheetBanco.getRange(Linha, 486).setValue(Entrada.getRange('B140').getValue());
      SheetBanco.getRange(Linha, 487).setValue(Entrada.getRange('B141').getValue());
      SheetBanco.getRange(Linha, 488).setValue(Entrada.getRange('B142').getValue());
      SheetBanco.getRange(Linha, 489).setValue(Entrada.getRange('B143').getValue());
      SheetBanco.getRange(Linha, 490).setValue(Entrada.getRange('B144').getValue());
      SheetBanco.getRange(Linha, 491).setValue(Entrada.getRange('B145').getValue());
      SheetBanco.getRange(Linha, 492).setValue(Entrada.getRange('B146').getValue());
      SheetBanco.getRange(Linha, 493).setValue(Entrada.getRange('B147').getValue());
      SheetBanco.getRange(Linha, 494).setValue(Entrada.getRange('B148').getValue());
      SheetBanco.getRange(Linha, 495).setValue(Entrada.getRange('B149').getValue());
      SheetBanco.getRange(Linha, 496).setValue(Entrada.getRange('B150').getValue());
      SheetBanco.getRange(Linha, 497).setValue(Entrada.getRange('B151').getValue());
      SheetBanco.getRange(Linha, 498).setValue(Entrada.getRange('B152').getValue());
      SheetBanco.getRange(Linha, 499).setValue(Entrada.getRange('B153').getValue());
      SheetBanco.getRange(Linha, 500).setValue(Entrada.getRange('B154').getValue());
      SheetBanco.getRange(Linha, 501).setValue(Entrada.getRange('B155').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('B156').getValue());
      SheetBanco.getRange(Linha, 502).setValue(Entrada.getRange('B157').getValue());
      SheetBanco.getRange(Linha, 503).setValue(Entrada.getRange('B158').getValue());
      SheetBanco.getRange(Linha, 504).setValue(Entrada.getRange('B159').getValue());
      SheetBanco.getRange(Linha, 505).setValue(Entrada.getRange('B160').getValue());
      SheetBanco.getRange(Linha, 506).setValue(Entrada.getRange('B161').getValue());
      SheetBanco.getRange(Linha, 507).setValue(Entrada.getRange('B162').getValue());
      SheetBanco.getRange(Linha, 508).setValue(Entrada.getRange('B163').getValue());
      SheetBanco.getRange(Linha, 509).setValue(Entrada.getRange('B164').getValue());
      SheetBanco.getRange(Linha, 510).setValue(Entrada.getRange('B165').getValue());
      SheetBanco.getRange(Linha, 511).setValue(Entrada.getRange('B166').getValue());
      SheetBanco.getRange(Linha, 512).setValue(Entrada.getRange('B167').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('B168').getValue());
      SheetBanco.getRange(Linha, 513).setValue(Entrada.getRange('B169').getValue());
      SheetBanco.getRange(Linha, 514).setValue(Entrada.getRange('B170').getValue());
      SheetBanco.getRange(Linha, 515).setValue(Entrada.getRange('B171').getValue());
      SheetBanco.getRange(Linha, 516).setValue(Entrada.getRange('B172').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('B173').getValue());
      SheetBanco.getRange(Linha, 517).setValue(Entrada.getRange('B174').getValue());
      SheetBanco.getRange(Linha, 518).setValue(Entrada.getRange('B175').getValue());
      SheetBanco.getRange(Linha, 519).setValue(Entrada.getRange('B176').getValue());
      SheetBanco.getRange(Linha, 520).setValue(Entrada.getRange('B177').getValue());
      SheetBanco.getRange(Linha, 521).setValue(Entrada.getRange('B178').getValue());
      SheetBanco.getRange(Linha, 522).setValue(Entrada.getRange('B179').getValue());
      SheetBanco.getRange(Linha, 523).setValue(Entrada.getRange('B180').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('B181').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('B182').getValue());
      SheetBanco.getRange(Linha, 524).setValue(Entrada.getRange('D137').getValue());
      SheetBanco.getRange(Linha, 525).setValue(Entrada.getRange('D138').getValue());
      SheetBanco.getRange(Linha, 526).setValue(Entrada.getRange('D139').getValue());
      SheetBanco.getRange(Linha, 527).setValue(Entrada.getRange('D140').getValue());
      SheetBanco.getRange(Linha, 528).setValue(Entrada.getRange('D141').getValue());
      SheetBanco.getRange(Linha, 529).setValue(Entrada.getRange('D142').getValue());
      SheetBanco.getRange(Linha, 530).setValue(Entrada.getRange('D143').getValue());
      SheetBanco.getRange(Linha, 531).setValue(Entrada.getRange('D144').getValue());
      SheetBanco.getRange(Linha, 532).setValue(Entrada.getRange('D145').getValue());
      SheetBanco.getRange(Linha, 533).setValue(Entrada.getRange('D146').getValue());
      SheetBanco.getRange(Linha, 534).setValue(Entrada.getRange('D147').getValue());
      SheetBanco.getRange(Linha, 535).setValue(Entrada.getRange('D148').getValue());
      SheetBanco.getRange(Linha, 536).setValue(Entrada.getRange('D149').getValue());
      SheetBanco.getRange(Linha, 537).setValue(Entrada.getRange('D150').getValue());
      SheetBanco.getRange(Linha, 538).setValue(Entrada.getRange('D151').getValue());
      SheetBanco.getRange(Linha, 539).setValue(Entrada.getRange('D152').getValue());
      SheetBanco.getRange(Linha, 540).setValue(Entrada.getRange('D153').getValue());
      SheetBanco.getRange(Linha, 541).setValue(Entrada.getRange('D154').getValue());
      SheetBanco.getRange(Linha, 542).setValue(Entrada.getRange('D155').getValue());
      SheetBanco.getRange(Linha, 543).setValue(Entrada.getRange('D156').getValue());
      SheetBanco.getRange(Linha, 544).setValue(Entrada.getRange('D157').getValue());
      SheetBanco.getRange(Linha, 545).setValue(Entrada.getRange('D158').getValue());
      SheetBanco.getRange(Linha, 546).setValue(Entrada.getRange('D159').getValue());
      SheetBanco.getRange(Linha, 547).setValue(Entrada.getRange('D160').getValue());
      SheetBanco.getRange(Linha, 548).setValue(Entrada.getRange('D161').getValue());
      SheetBanco.getRange(Linha, 549).setValue(Entrada.getRange('D162').getValue());
      SheetBanco.getRange(Linha, 550).setValue(Entrada.getRange('D163').getValue());
      SheetBanco.getRange(Linha, 551).setValue(Entrada.getRange('D164').getValue());
      SheetBanco.getRange(Linha, 552).setValue(Entrada.getRange('D165').getValue());
      SheetBanco.getRange(Linha, 553).setValue(Entrada.getRange('D166').getValue());
      SheetBanco.getRange(Linha, 554).setValue(Entrada.getRange('D167').getValue());
      SheetBanco.getRange(Linha, 555).setValue(Entrada.getRange('D168').getValue());
      SheetBanco.getRange(Linha, 556).setValue(Entrada.getRange('D169').getValue());
      SheetBanco.getRange(Linha, 557).setValue(Entrada.getRange('D170').getValue());
      SheetBanco.getRange(Linha, 558).setValue(Entrada.getRange('D171').getValue());
      SheetBanco.getRange(Linha, 559).setValue(Entrada.getRange('D172').getValue());
      SheetBanco.getRange(Linha, 560).setValue(Entrada.getRange('D173').getValue());
      SheetBanco.getRange(Linha, 561).setValue(Entrada.getRange('D174').getValue());
      SheetBanco.getRange(Linha, 562).setValue(Entrada.getRange('D175').getValue());
      SheetBanco.getRange(Linha, 563).setValue(Entrada.getRange('D176').getValue());
      SheetBanco.getRange(Linha, 564).setValue(Entrada.getRange('D177').getValue());
      SheetBanco.getRange(Linha, 565).setValue(Entrada.getRange('D178').getValue());
      SheetBanco.getRange(Linha, 566).setValue(Entrada.getRange('D179').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('D180').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('D181').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('D182').getValue());
      SheetBanco.getRange(Linha, 567).setValue(Entrada.getRange('F137').getValue());
      SheetBanco.getRange(Linha, 568).setValue(Entrada.getRange('F138').getValue());
      SheetBanco.getRange(Linha, 569).setValue(Entrada.getRange('F139').getValue());
      SheetBanco.getRange(Linha, 570).setValue(Entrada.getRange('F140').getValue());
      SheetBanco.getRange(Linha, 571).setValue(Entrada.getRange('F141').getValue());
      SheetBanco.getRange(Linha, 572).setValue(Entrada.getRange('F142').getValue());
      SheetBanco.getRange(Linha, 573).setValue(Entrada.getRange('F143').getValue());
      SheetBanco.getRange(Linha, 574).setValue(Entrada.getRange('F144').getValue());
      SheetBanco.getRange(Linha, 575).setValue(Entrada.getRange('F145').getValue());
      SheetBanco.getRange(Linha, 576).setValue(Entrada.getRange('F146').getValue());
      SheetBanco.getRange(Linha, 577).setValue(Entrada.getRange('F147').getValue());
      SheetBanco.getRange(Linha, 578).setValue(Entrada.getRange('F148').getValue());
      SheetBanco.getRange(Linha, 579).setValue(Entrada.getRange('F149').getValue());
      SheetBanco.getRange(Linha, 580).setValue(Entrada.getRange('F150').getValue());
      SheetBanco.getRange(Linha, 581).setValue(Entrada.getRange('F151').getValue());
      SheetBanco.getRange(Linha, 582).setValue(Entrada.getRange('F152').getValue());
      SheetBanco.getRange(Linha, 583).setValue(Entrada.getRange('F153').getValue());
      SheetBanco.getRange(Linha, 584).setValue(Entrada.getRange('F154').getValue());
      SheetBanco.getRange(Linha, 585).setValue(Entrada.getRange('F155').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('F156').getValue());
      SheetBanco.getRange(Linha, 586).setValue(Entrada.getRange('F157').getValue());
      SheetBanco.getRange(Linha, 587).setValue(Entrada.getRange('F158').getValue());
      SheetBanco.getRange(Linha, 588).setValue(Entrada.getRange('F159').getValue());
      SheetBanco.getRange(Linha, 589).setValue(Entrada.getRange('F160').getValue());
      SheetBanco.getRange(Linha, 590).setValue(Entrada.getRange('F161').getValue());
      SheetBanco.getRange(Linha, 591).setValue(Entrada.getRange('F162').getValue());
      SheetBanco.getRange(Linha, 592).setValue(Entrada.getRange('F163').getValue());
      SheetBanco.getRange(Linha, 593).setValue(Entrada.getRange('F164').getValue());
      SheetBanco.getRange(Linha, 594).setValue(Entrada.getRange('F165').getValue());
      SheetBanco.getRange(Linha, 595).setValue(Entrada.getRange('F166').getValue());
      SheetBanco.getRange(Linha, 596).setValue(Entrada.getRange('F167').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('F168').getValue());
      SheetBanco.getRange(Linha, 597).setValue(Entrada.getRange('F169').getValue());
      SheetBanco.getRange(Linha, 598).setValue(Entrada.getRange('F170').getValue());
      SheetBanco.getRange(Linha, 599).setValue(Entrada.getRange('F171').getValue());
      SheetBanco.getRange(Linha, 600).setValue(Entrada.getRange('F172').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('F173').getValue());
      SheetBanco.getRange(Linha, 601).setValue(Entrada.getRange('F174').getValue());
      SheetBanco.getRange(Linha, 602).setValue(Entrada.getRange('F175').getValue());
      SheetBanco.getRange(Linha, 603).setValue(Entrada.getRange('F176').getValue());
      SheetBanco.getRange(Linha, 604).setValue(Entrada.getRange('F177').getValue());
      SheetBanco.getRange(Linha, 605).setValue(Entrada.getRange('F178').getValue());
      SheetBanco.getRange(Linha, 606).setValue(Entrada.getRange('F179').getValue());
      SheetBanco.getRange(Linha, 607).setValue(Entrada.getRange('F180').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('F181').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('F182').getValue());
      SheetBanco.getRange(Linha, 608).setValue(Entrada.getRange('H137').getValue());
      SheetBanco.getRange(Linha, 609).setValue(Entrada.getRange('H138').getValue());
      SheetBanco.getRange(Linha, 610).setValue(Entrada.getRange('H139').getValue());
      SheetBanco.getRange(Linha, 611).setValue(Entrada.getRange('H140').getValue());
      SheetBanco.getRange(Linha, 612).setValue(Entrada.getRange('H141').getValue());
      SheetBanco.getRange(Linha, 613).setValue(Entrada.getRange('H142').getValue());
      SheetBanco.getRange(Linha, 614).setValue(Entrada.getRange('H143').getValue());
      SheetBanco.getRange(Linha, 615).setValue(Entrada.getRange('H144').getValue());
      SheetBanco.getRange(Linha, 616).setValue(Entrada.getRange('H145').getValue());
      SheetBanco.getRange(Linha, 617).setValue(Entrada.getRange('H146').getValue());
      SheetBanco.getRange(Linha, 618).setValue(Entrada.getRange('H147').getValue());
      SheetBanco.getRange(Linha, 619).setValue(Entrada.getRange('H148').getValue());
      SheetBanco.getRange(Linha, 620).setValue(Entrada.getRange('H149').getValue());
      SheetBanco.getRange(Linha, 621).setValue(Entrada.getRange('H150').getValue());
      SheetBanco.getRange(Linha, 622).setValue(Entrada.getRange('H151').getValue());
      SheetBanco.getRange(Linha, 623).setValue(Entrada.getRange('H152').getValue());
      SheetBanco.getRange(Linha, 624).setValue(Entrada.getRange('H153').getValue());
      SheetBanco.getRange(Linha, 625).setValue(Entrada.getRange('H154').getValue());
      SheetBanco.getRange(Linha, 626).setValue(Entrada.getRange('H155').getValue());
      SheetBanco.getRange(Linha, 627).setValue(Entrada.getRange('H156').getValue());
      SheetBanco.getRange(Linha, 628).setValue(Entrada.getRange('H157').getValue());
      SheetBanco.getRange(Linha, 629).setValue(Entrada.getRange('H158').getValue());
      SheetBanco.getRange(Linha, 630).setValue(Entrada.getRange('H159').getValue());
      SheetBanco.getRange(Linha, 631).setValue(Entrada.getRange('H160').getValue());
      SheetBanco.getRange(Linha, 632).setValue(Entrada.getRange('H161').getValue());
      SheetBanco.getRange(Linha, 633).setValue(Entrada.getRange('H162').getValue());
      SheetBanco.getRange(Linha, 634).setValue(Entrada.getRange('H163').getValue());
      SheetBanco.getRange(Linha, 635).setValue(Entrada.getRange('H164').getValue());
      SheetBanco.getRange(Linha, 636).setValue(Entrada.getRange('H165').getValue());
      SheetBanco.getRange(Linha, 637).setValue(Entrada.getRange('H166').getValue());
      SheetBanco.getRange(Linha, 638).setValue(Entrada.getRange('H167').getValue());
      SheetBanco.getRange(Linha, 639).setValue(Entrada.getRange('H168').getValue());
      SheetBanco.getRange(Linha, 640).setValue(Entrada.getRange('H169').getValue());
      SheetBanco.getRange(Linha, 641).setValue(Entrada.getRange('H170').getValue());
      SheetBanco.getRange(Linha, 642).setValue(Entrada.getRange('H171').getValue());
      SheetBanco.getRange(Linha, 643).setValue(Entrada.getRange('H172').getValue());
      SheetBanco.getRange(Linha, 644).setValue(Entrada.getRange('H173').getValue());
      SheetBanco.getRange(Linha, 645).setValue(Entrada.getRange('H174').getValue());
      SheetBanco.getRange(Linha, 646).setValue(Entrada.getRange('H175').getValue());
      SheetBanco.getRange(Linha, 647).setValue(Entrada.getRange('H176').getValue());
      SheetBanco.getRange(Linha, 648).setValue(Entrada.getRange('H177').getValue());
      SheetBanco.getRange(Linha, 649).setValue(Entrada.getRange('H178').getValue());
      SheetBanco.getRange(Linha, 650).setValue(Entrada.getRange('H179').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('H180').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('H181').getValue());
      //SheetBanco.getRange(Linha, col).setValue(Entrada.getRange('H182').getValue());
    
      LimparEntrada();
    
  
};

function ExcluirOcorrencia(){
  
  var Entrada = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Entrada'); 
  var Ocorrencia = Entrada.getRange('B2').getValue();
  
 
  var SheetBanco = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Banco');
    
  SheetBanco.getActiveCell();
  
  var LocalPesquisa = SheetBanco.getRange(2,2, SheetBanco.getLastRow()).getValues();
  var Resultado = LocalPesquisa.Pesquisa(Ocorrencia);
  
  var Linha1 = Resultado + 2
  
  
  if(Resultado != -1){
    
    LocalizarCelulaVazia_Excluida()
  
    var Excluida = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Excluidas');
    var Linha2 = Excluida.getCurrentCell().getRow();
       
      
    Excluida.getRange(Linha2, 2).setValue(SheetBanco.getRange(Linha1, 2).getValue());
    Excluida.getRange(Linha2, 3).setValue(SheetBanco.getRange(Linha1, 3).getValue());
    Excluida.getRange(Linha2, 4).setValue(SheetBanco.getRange(Linha1, 4).getValue());
    Excluida.getRange(Linha2, 5).setValue(SheetBanco.getRange(Linha1, 6).getValue());
    Excluida.getRange(Linha2, 6).setValue(SheetBanco.getRange(Linha1, 11).getValue());
    Excluida.getRange(Linha2, 7).setValue(SheetBanco.getRange(Linha1, 13).getValue());
    Excluida.getRange(Linha2, 8).setValue(SheetBanco.getRange(Linha1, 23).getValue());
    Excluida.getRange(Linha2, 9).setValue(SheetBanco.getRange(Linha1, 28).getValue());
    Excluida.getRange(Linha2, 10).setValue(SheetBanco.getRange(Linha1, 40).getValue());
    Excluida.getRange(Linha2, 11).setValue(SheetBanco.getRange(Linha1, 41).getValue());
    Excluida.getRange(Linha2, 12).setValue(SheetBanco.getRange(Linha1, 44).getValue());
    Excluida.getRange(Linha2, 13).setValue(SheetBanco.getRange(Linha1, 47).getValue());
    Excluida.getRange(Linha2, 14).setValue(SheetBanco.getRange(Linha1, 66).getValue());
    Excluida.getRange(Linha2, 15).setValue(SheetBanco.getRange(Linha1, 69).getValue());
    Excluida.getRange(Linha2, 16).setValue(SheetBanco.getRange(Linha1, 72).getValue());
    Excluida.getRange(Linha2, 17).setValue(SheetBanco.getRange(Linha1, 78).getValue());
    Excluida.getRange(Linha2, 18).setValue(SheetBanco.getRange(Linha1, 39).getValue());
        
        
    SheetBanco.deleteRow(Linha1);
    
     
    Browser.msgBox("Ocorrência excluída!")
  
    LimparEntrada();
  }else{
    
    Browser.msgBox("Ocorrência não localizada!")
  
  }
  
};

Array.prototype.Pesquisa = function(Procura){
  if(Procura == "") return false;
  for(var Linha = 0; Linha < this.length;Linha++)
    if(this[Linha] == Procura) return Linha; 
   return -1

}

function addNA(){
	var Entrada = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Entrada');
	var acusados = Entrada.getRange(2,4).getValue();

	if(acusados == 0){
	//ACUSADO 1
    Entrada.getRange('B43').setValue('NA');
    Entrada.getRange('B44').setValue('NA');
    Entrada.getRange('B45').setValue('NA');
    Entrada.getRange('B46').setValue('NA');
    Entrada.getRange('B47').setValue('NA');
    Entrada.getRange('B51').setValue('NA');
    Entrada.getRange('B52').setValue('NA');
    Entrada.getRange('B53').setValue('NA');
    Entrada.getRange('B54').setValue('NA');
    Entrada.getRange('B55').setValue('NA');
    Entrada.getRange('B56').setValue('NA');
    Entrada.getRange('B57').setValue('NA');
    Entrada.getRange('B58').setValue('NA');
    Entrada.getRange('B59').setValue('NA');
    Entrada.getRange('B60').setValue('NA');
    Entrada.getRange('B61').setValue('NA');
    //Entrada.getRange('B62').setValue('NA');
    Entrada.getRange('B63').setValue('NA');
    Entrada.getRange('B64').setValue('NA');
    Entrada.getRange('B65').setValue('NA');
    Entrada.getRange('B66').setValue('NA');
    Entrada.getRange('B67').setValue('NA');
    Entrada.getRange('B68').setValue('NA');
    Entrada.getRange('B69').setValue('NA');
    Entrada.getRange('B70').setValue('NA');
    Entrada.getRange('B71').setValue('NA');
    Entrada.getRange('B72').setValue('NA');
    Entrada.getRange('B73').setValue('NA');
    //Entrada.getRange('B74').setValue('NA');
    Entrada.getRange('B75').setValue('NA');
    Entrada.getRange('B76').setValue('NA');
    Entrada.getRange('B77').setValue('NA');
    Entrada.getRange('B78').setValue('NA');
    //Entrada.getRange('B79').setValue('NA');
    Entrada.getRange('B80').setValue('NA');
    Entrada.getRange('B81').setValue('NA');
    Entrada.getRange('B82').setValue('NA');
    Entrada.getRange('B83').setValue('NA');
    Entrada.getRange('B84').setValue('NA');
    Entrada.getRange('B85').setValue('NA');
    Entrada.getRange('B86').setValue('NA');
    //Entrada.getRange('B87').setValue('NA');
    //Entrada.getRange('B88').setValue('NA');
    Entrada.getRange('D43').setValue('NA');
    Entrada.getRange('D44').setValue('NA');
    Entrada.getRange('D45').setValue('NA');
    Entrada.getRange('D46').setValue('NA');
    Entrada.getRange('D47').setValue('NA');
    Entrada.getRange('D48').setValue('NA');
    Entrada.getRange('D49').setValue('NA');
    Entrada.getRange('D50').setValue('NA');
    Entrada.getRange('D51').setValue('NA');
    Entrada.getRange('D52').setValue('NA');
    Entrada.getRange('D53').setValue('NA');
    Entrada.getRange('D54').setValue('NA');
    Entrada.getRange('D55').setValue('NA');
    Entrada.getRange('D56').setValue('NA');
    Entrada.getRange('D57').setValue('NA');
    Entrada.getRange('D58').setValue('NA');
    Entrada.getRange('D59').setValue('NA');
    Entrada.getRange('D60').setValue('NA');
    Entrada.getRange('D61').setValue('NA');
    Entrada.getRange('D62').setValue('NA');
    Entrada.getRange('D63').setValue('NA');
    Entrada.getRange('D64').setValue('NA');
    Entrada.getRange('D65').setValue('NA');
    Entrada.getRange('D66').setValue('NA');
    Entrada.getRange('D67').setValue('NA');
    Entrada.getRange('D68').setValue('NA');
    Entrada.getRange('D69').setValue('NA');
    Entrada.getRange('D70').setValue('NA');
    Entrada.getRange('D71').setValue('NA');
    Entrada.getRange('D72').setValue('NA');
    Entrada.getRange('D73').setValue('NA');
    Entrada.getRange('D74').setValue('NA');
    Entrada.getRange('D75').setValue('NA');
    Entrada.getRange('D76').setValue('NA');
    Entrada.getRange('D77').setValue('NA');
    Entrada.getRange('D78').setValue('NA');
    Entrada.getRange('D79').setValue('NA');
    Entrada.getRange('D80').setValue('NA');
    Entrada.getRange('D81').setValue('NA');
    Entrada.getRange('D82').setValue('NA');
    Entrada.getRange('D83').setValue('NA');
    Entrada.getRange('D84').setValue('NA');
    Entrada.getRange('D85').setValue('NA');
    //Entrada.getRange('D86').setValue('NA');
    //Entrada.getRange('D87').setValue('NA');
    //Entrada.getRange('D88').setValue('NA');
	//ACUSADO 2     
    Entrada.getRange('F43').setValue('NA');
    Entrada.getRange('F44').setValue('NA');
    Entrada.getRange('F45').setValue('NA');
    Entrada.getRange('F46').setValue('NA');
    Entrada.getRange('F47').setValue('NA');
    Entrada.getRange('F51').setValue('NA');
    Entrada.getRange('F52').setValue('NA');
    Entrada.getRange('F53').setValue('NA');
    Entrada.getRange('F54').setValue('NA');
    Entrada.getRange('F55').setValue('NA');
    Entrada.getRange('F56').setValue('NA');
    Entrada.getRange('F57').setValue('NA');
    Entrada.getRange('F58').setValue('NA');
    Entrada.getRange('F59').setValue('NA');
    Entrada.getRange('F60').setValue('NA');
    Entrada.getRange('F61').setValue('NA');
    //Entrada.getRange('F62').setValue('NA');
    Entrada.getRange('F63').setValue('NA');
    Entrada.getRange('F64').setValue('NA');
    Entrada.getRange('F65').setValue('NA');
    Entrada.getRange('F66').setValue('NA');
    Entrada.getRange('F67').setValue('NA');
    Entrada.getRange('F68').setValue('NA');
    Entrada.getRange('F69').setValue('NA');
    Entrada.getRange('F70').setValue('NA');
    Entrada.getRange('F71').setValue('NA');
    Entrada.getRange('F72').setValue('NA');
    Entrada.getRange('F73').setValue('NA');
    //Entrada.getRange('F74').setValue('NA');
    Entrada.getRange('F75').setValue('NA');
    Entrada.getRange('F76').setValue('NA');
    Entrada.getRange('F77').setValue('NA');
    Entrada.getRange('F78').setValue('NA');
    //Entrada.getRange('F79').setValue('NA');
    Entrada.getRange('F80').setValue('NA');
    Entrada.getRange('F81').setValue('NA');
    Entrada.getRange('F82').setValue('NA');
    Entrada.getRange('F83').setValue('NA');
    Entrada.getRange('F84').setValue('NA');
    Entrada.getRange('F85').setValue('NA');
    Entrada.getRange('F86').setValue('NA');
    //Entrada.getRange('F87').setValue('NA');
    //Entrada.getRange('F88').setValue('NA');
    Entrada.getRange('H43').setValue('NA');
    Entrada.getRange('H44').setValue('NA');
    Entrada.getRange('H45').setValue('NA');
    Entrada.getRange('H46').setValue('NA');
    Entrada.getRange('H47').setValue('NA');
    Entrada.getRange('H48').setValue('NA');
    Entrada.getRange('H49').setValue('NA');
    Entrada.getRange('H50').setValue('NA');
    Entrada.getRange('H51').setValue('NA');
    Entrada.getRange('H52').setValue('NA');
    Entrada.getRange('H53').setValue('NA');
    Entrada.getRange('H54').setValue('NA');
    Entrada.getRange('H55').setValue('NA');
    Entrada.getRange('H56').setValue('NA');
    Entrada.getRange('H57').setValue('NA');
    Entrada.getRange('H58').setValue('NA');
    Entrada.getRange('H59').setValue('NA');
    Entrada.getRange('H60').setValue('NA');
    Entrada.getRange('H61').setValue('NA');
    Entrada.getRange('H62').setValue('NA');
    Entrada.getRange('H63').setValue('NA');
    Entrada.getRange('H64').setValue('NA');
    Entrada.getRange('H65').setValue('NA');
    Entrada.getRange('H66').setValue('NA');
    Entrada.getRange('H67').setValue('NA');
    Entrada.getRange('H68').setValue('NA');
    Entrada.getRange('H69').setValue('NA');
    Entrada.getRange('H70').setValue('NA');
    Entrada.getRange('H71').setValue('NA');
    Entrada.getRange('H72').setValue('NA');
    Entrada.getRange('H73').setValue('NA');
    Entrada.getRange('H74').setValue('NA');
    Entrada.getRange('H75').setValue('NA');
    Entrada.getRange('H76').setValue('NA');
    Entrada.getRange('H77').setValue('NA');
    Entrada.getRange('H78').setValue('NA');
    Entrada.getRange('H79').setValue('NA');
    Entrada.getRange('H80').setValue('NA');
    Entrada.getRange('H81').setValue('NA');
    Entrada.getRange('H82').setValue('NA');
    Entrada.getRange('H83').setValue('NA');
    Entrada.getRange('H84').setValue('NA');
    Entrada.getRange('H85').setValue('NA');
    //Entrada.getRange('H86').setValue('NA');
    //Entrada.getRange('H87').setValue('NA');
    //Entrada.getRange('H88').setValue('NA');
	//ACUSADO 3
    Entrada.getRange('B90').setValue('NA');
    Entrada.getRange('B91').setValue('NA');
    Entrada.getRange('B92').setValue('NA');
    Entrada.getRange('B93').setValue('NA');
    Entrada.getRange('B94').setValue('NA');
    Entrada.getRange('B98').setValue('NA');
    Entrada.getRange('B99').setValue('NA');
    Entrada.getRange('B100').setValue('NA');
    Entrada.getRange('B101').setValue('NA');
    Entrada.getRange('B102').setValue('NA');
    Entrada.getRange('B103').setValue('NA');
    Entrada.getRange('B104').setValue('NA');
    Entrada.getRange('B105').setValue('NA');
    Entrada.getRange('B106').setValue('NA');
    Entrada.getRange('B107').setValue('NA');
    Entrada.getRange('B108').setValue('NA');
    //Entrada.getRange('B109').setValue('NA');
    Entrada.getRange('B110').setValue('NA');
    Entrada.getRange('B111').setValue('NA');
    Entrada.getRange('B112').setValue('NA');
    Entrada.getRange('B113').setValue('NA');
    Entrada.getRange('B114').setValue('NA');
    Entrada.getRange('B115').setValue('NA');
    Entrada.getRange('B116').setValue('NA');
    Entrada.getRange('B117').setValue('NA');
    Entrada.getRange('B118').setValue('NA');
    Entrada.getRange('B119').setValue('NA');
    Entrada.getRange('B120').setValue('NA');
    //Entrada.getRange('B121').setValue('NA');
    Entrada.getRange('B122').setValue('NA');
    Entrada.getRange('B123').setValue('NA');
    Entrada.getRange('B124').setValue('NA');
    Entrada.getRange('B125').setValue('NA');
    //Entrada.getRange('B126').setValue('NA');
    Entrada.getRange('B127').setValue('NA');
    Entrada.getRange('B128').setValue('NA');
    Entrada.getRange('B129').setValue('NA');
    Entrada.getRange('B130').setValue('NA');
    Entrada.getRange('B131').setValue('NA');
    Entrada.getRange('B132').setValue('NA');
    Entrada.getRange('B133').setValue('NA');
    //Entrada.getRange('B134').setValue('NA');
    //Entrada.getRange('B135').setValue('NA');
    Entrada.getRange('D90').setValue('NA');
    Entrada.getRange('D91').setValue('NA');
    Entrada.getRange('D92').setValue('NA');
    Entrada.getRange('D93').setValue('NA');
    Entrada.getRange('D94').setValue('NA');
    Entrada.getRange('D95').setValue('NA');
    Entrada.getRange('D96').setValue('NA');
    Entrada.getRange('D97').setValue('NA');
    Entrada.getRange('D98').setValue('NA');
    Entrada.getRange('D99').setValue('NA');
    Entrada.getRange('D100').setValue('NA');
    Entrada.getRange('D101').setValue('NA');
    Entrada.getRange('D102').setValue('NA');
    Entrada.getRange('D103').setValue('NA');
    Entrada.getRange('D104').setValue('NA');
    Entrada.getRange('D105').setValue('NA');
    Entrada.getRange('D106').setValue('NA');
    Entrada.getRange('D107').setValue('NA');
    Entrada.getRange('D108').setValue('NA');
    Entrada.getRange('D109').setValue('NA');
    Entrada.getRange('D110').setValue('NA');
    Entrada.getRange('D111').setValue('NA');
    Entrada.getRange('D112').setValue('NA');
    Entrada.getRange('D113').setValue('NA');
    Entrada.getRange('D114').setValue('NA');
    Entrada.getRange('D115').setValue('NA');
    Entrada.getRange('D116').setValue('NA');
    Entrada.getRange('D117').setValue('NA');
    Entrada.getRange('D118').setValue('NA');
    Entrada.getRange('D119').setValue('NA');
    Entrada.getRange('D120').setValue('NA');
    Entrada.getRange('D121').setValue('NA');
    Entrada.getRange('D122').setValue('NA');
    Entrada.getRange('D123').setValue('NA');
    Entrada.getRange('D124').setValue('NA');
    Entrada.getRange('D125').setValue('NA');
    Entrada.getRange('D126').setValue('NA');
    Entrada.getRange('D127').setValue('NA');
    Entrada.getRange('D128').setValue('NA');
    Entrada.getRange('D129').setValue('NA');
    Entrada.getRange('D130').setValue('NA');
    Entrada.getRange('D131').setValue('NA');
    Entrada.getRange('D132').setValue('NA');
    //Entrada.getRange('D133').setValue('NA');
    //Entrada.getRange('D134').setValue('NA');
    //Entrada.getRange('D135').setValue('NA');
	//ACUSADO 4
    Entrada.getRange('F90').setValue('NA');
    Entrada.getRange('F91').setValue('NA');
    Entrada.getRange('F92').setValue('NA');
    Entrada.getRange('F93').setValue('NA');
    Entrada.getRange('F94').setValue('NA');
    Entrada.getRange('F98').setValue('NA');
    Entrada.getRange('F99').setValue('NA');
    Entrada.getRange('F100').setValue('NA');
    Entrada.getRange('F101').setValue('NA');
    Entrada.getRange('F102').setValue('NA');
    Entrada.getRange('F103').setValue('NA');
    Entrada.getRange('F104').setValue('NA');
    Entrada.getRange('F105').setValue('NA');
    Entrada.getRange('F106').setValue('NA');
    Entrada.getRange('F107').setValue('NA');
    Entrada.getRange('F108').setValue('NA');
    //Entrada.getRange('F109').setValue('NA');
    Entrada.getRange('F110').setValue('NA');
    Entrada.getRange('F111').setValue('NA');
    Entrada.getRange('F112').setValue('NA');
    Entrada.getRange('F113').setValue('NA');
    Entrada.getRange('F114').setValue('NA');
    Entrada.getRange('F115').setValue('NA');
    Entrada.getRange('F116').setValue('NA');
    Entrada.getRange('F117').setValue('NA');
    Entrada.getRange('F118').setValue('NA');
    Entrada.getRange('F119').setValue('NA');
    Entrada.getRange('F120').setValue('NA');
    //Entrada.getRange('F121').setValue('NA');
    Entrada.getRange('F122').setValue('NA');
    Entrada.getRange('F123').setValue('NA');
    Entrada.getRange('F124').setValue('NA');
    Entrada.getRange('F125').setValue('NA');
    //Entrada.getRange('F126').setValue('NA');
    Entrada.getRange('F127').setValue('NA');
    Entrada.getRange('F128').setValue('NA');
    Entrada.getRange('F129').setValue('NA');
    Entrada.getRange('F130').setValue('NA');
    Entrada.getRange('F131').setValue('NA');
    Entrada.getRange('F132').setValue('NA');
    Entrada.getRange('F133').setValue('NA');
    //Entrada.getRange('F134').setValue('NA');
    //Entrada.getRange('F135').setValue('NA');
    Entrada.getRange('H90').setValue('NA');
    Entrada.getRange('H91').setValue('NA');
    Entrada.getRange('H92').setValue('NA');
    Entrada.getRange('H93').setValue('NA');
    Entrada.getRange('H94').setValue('NA');
    Entrada.getRange('H95').setValue('NA');
    Entrada.getRange('H96').setValue('NA');
    Entrada.getRange('H97').setValue('NA');
    Entrada.getRange('H98').setValue('NA');
    Entrada.getRange('H99').setValue('NA');
    Entrada.getRange('H100').setValue('NA');
    Entrada.getRange('H101').setValue('NA');
    Entrada.getRange('H102').setValue('NA');
    Entrada.getRange('H103').setValue('NA');
    Entrada.getRange('H104').setValue('NA');
    Entrada.getRange('H105').setValue('NA');
    Entrada.getRange('H106').setValue('NA');
    Entrada.getRange('H107').setValue('NA');
    Entrada.getRange('H108').setValue('NA');
    Entrada.getRange('H109').setValue('NA');
    Entrada.getRange('H110').setValue('NA');
    Entrada.getRange('H111').setValue('NA');
    Entrada.getRange('H112').setValue('NA');
    Entrada.getRange('H113').setValue('NA');
    Entrada.getRange('H114').setValue('NA');
    Entrada.getRange('H115').setValue('NA');
    Entrada.getRange('H116').setValue('NA');
    Entrada.getRange('H117').setValue('NA');
    Entrada.getRange('H118').setValue('NA');
    Entrada.getRange('H119').setValue('NA');
    Entrada.getRange('H120').setValue('NA');
    Entrada.getRange('H121').setValue('NA');
    Entrada.getRange('H122').setValue('NA');
    Entrada.getRange('H123').setValue('NA');
    Entrada.getRange('H124').setValue('NA');
    Entrada.getRange('H125').setValue('NA');
    Entrada.getRange('H126').setValue('NA');
    Entrada.getRange('H127').setValue('NA');
    Entrada.getRange('H128').setValue('NA');
    Entrada.getRange('H129').setValue('NA');
    Entrada.getRange('H130').setValue('NA');
    Entrada.getRange('H131').setValue('NA');
    Entrada.getRange('H132').setValue('NA');
    //Entrada.getRange('H133').setValue('NA');
    //Entrada.getRange('H134').setValue('NA');
    //Entrada.getRange('H135').setValue('NA');
	//ACUSADO 5     
    Entrada.getRange('B137').setValue('NA');
    Entrada.getRange('B138').setValue('NA');
    Entrada.getRange('B139').setValue('NA');
    Entrada.getRange('B140').setValue('NA');
    Entrada.getRange('B141').setValue('NA');
    Entrada.getRange('B145').setValue('NA');
    Entrada.getRange('B146').setValue('NA');
    Entrada.getRange('B147').setValue('NA');
    Entrada.getRange('B148').setValue('NA');
    Entrada.getRange('B149').setValue('NA');
    Entrada.getRange('B150').setValue('NA');
    Entrada.getRange('B151').setValue('NA');
    Entrada.getRange('B152').setValue('NA');
    Entrada.getRange('B153').setValue('NA');
    Entrada.getRange('B154').setValue('NA');
    Entrada.getRange('B155').setValue('NA');
    //Entrada.getRange('B156').setValue('NA');
    Entrada.getRange('B157').setValue('NA');
    Entrada.getRange('B158').setValue('NA');
    Entrada.getRange('B159').setValue('NA');
    Entrada.getRange('B160').setValue('NA');
    Entrada.getRange('B161').setValue('NA');
    Entrada.getRange('B162').setValue('NA');
    Entrada.getRange('B163').setValue('NA');
    Entrada.getRange('B164').setValue('NA');
    Entrada.getRange('B165').setValue('NA');
    Entrada.getRange('B166').setValue('NA');
    Entrada.getRange('B167').setValue('NA');
    //Entrada.getRange('B168').setValue('NA');
    Entrada.getRange('B169').setValue('NA');
    Entrada.getRange('B170').setValue('NA');
    Entrada.getRange('B171').setValue('NA');
    Entrada.getRange('B172').setValue('NA');
    //Entrada.getRange('B173').setValue('NA');
    Entrada.getRange('B174').setValue('NA');
    Entrada.getRange('B175').setValue('NA');
    Entrada.getRange('B176').setValue('NA');
    Entrada.getRange('B177').setValue('NA');
    Entrada.getRange('B178').setValue('NA');
    Entrada.getRange('B179').setValue('NA');
    Entrada.getRange('B180').setValue('NA');
    //Entrada.getRange('B181').setValue('NA');
    //Entrada.getRange('B182').setValue('NA');
    Entrada.getRange('D137').setValue('NA');
    Entrada.getRange('D138').setValue('NA');
    Entrada.getRange('D139').setValue('NA');
    Entrada.getRange('D140').setValue('NA');
    Entrada.getRange('D141').setValue('NA');
    Entrada.getRange('D142').setValue('NA');
    Entrada.getRange('D143').setValue('NA');
    Entrada.getRange('D144').setValue('NA');
    Entrada.getRange('D145').setValue('NA');
    Entrada.getRange('D146').setValue('NA');
    Entrada.getRange('D147').setValue('NA');
    Entrada.getRange('D148').setValue('NA');
    Entrada.getRange('D149').setValue('NA');
    Entrada.getRange('D150').setValue('NA');
    Entrada.getRange('D151').setValue('NA');
    Entrada.getRange('D152').setValue('NA');
    Entrada.getRange('D153').setValue('NA');
    Entrada.getRange('D154').setValue('NA');
    Entrada.getRange('D155').setValue('NA');
    Entrada.getRange('D156').setValue('NA');
    Entrada.getRange('D157').setValue('NA');
    Entrada.getRange('D158').setValue('NA');
    Entrada.getRange('D159').setValue('NA');
    Entrada.getRange('D160').setValue('NA');
    Entrada.getRange('D161').setValue('NA');
    Entrada.getRange('D162').setValue('NA');
    Entrada.getRange('D163').setValue('NA');
    Entrada.getRange('D164').setValue('NA');
    Entrada.getRange('D165').setValue('NA');
    Entrada.getRange('D166').setValue('NA');
    Entrada.getRange('D167').setValue('NA');
    Entrada.getRange('D168').setValue('NA');
    Entrada.getRange('D169').setValue('NA');
    Entrada.getRange('D170').setValue('NA');
    Entrada.getRange('D171').setValue('NA');
    Entrada.getRange('D172').setValue('NA');
    Entrada.getRange('D173').setValue('NA');
    Entrada.getRange('D174').setValue('NA');
    Entrada.getRange('D175').setValue('NA');
    Entrada.getRange('D176').setValue('NA');
    Entrada.getRange('D177').setValue('NA');
    Entrada.getRange('D178').setValue('NA');
    Entrada.getRange('D179').setValue('NA');
    //Entrada.getRange('D180').setValue('NA');
    //Entrada.getRange('D181').setValue('NA');
    //Entrada.getRange('D182').setValue('NA');
	//ACUSADO 6     
    Entrada.getRange('F137').setValue('NA');
    Entrada.getRange('F138').setValue('NA');
    Entrada.getRange('F139').setValue('NA');
    Entrada.getRange('F140').setValue('NA');
    Entrada.getRange('F141').setValue('NA');
    Entrada.getRange('F145').setValue('NA');
    Entrada.getRange('F146').setValue('NA');
    Entrada.getRange('F147').setValue('NA');
    Entrada.getRange('F148').setValue('NA');
    Entrada.getRange('F149').setValue('NA');
    Entrada.getRange('F150').setValue('NA');
    Entrada.getRange('F151').setValue('NA');
    Entrada.getRange('F152').setValue('NA');
    Entrada.getRange('F153').setValue('NA');
    Entrada.getRange('F154').setValue('NA');
    Entrada.getRange('F155').setValue('NA');
    //Entrada.getRange('F156').setValue('NA');
    Entrada.getRange('F157').setValue('NA');
    Entrada.getRange('F158').setValue('NA');
    Entrada.getRange('F159').setValue('NA');
    Entrada.getRange('F160').setValue('NA');
    Entrada.getRange('F161').setValue('NA');
    Entrada.getRange('F162').setValue('NA');
    Entrada.getRange('F163').setValue('NA');
    Entrada.getRange('F164').setValue('NA');
    Entrada.getRange('F165').setValue('NA');
    Entrada.getRange('F166').setValue('NA');
    Entrada.getRange('F167').setValue('NA');
    //Entrada.getRange('F168').setValue('NA');
    Entrada.getRange('F169').setValue('NA');
    Entrada.getRange('F170').setValue('NA');
    Entrada.getRange('F171').setValue('NA');
    Entrada.getRange('F172').setValue('NA');
    //Entrada.getRange('F173').setValue('NA');
    Entrada.getRange('F174').setValue('NA');
    Entrada.getRange('F175').setValue('NA');
    Entrada.getRange('F176').setValue('NA');
    Entrada.getRange('F177').setValue('NA');
    Entrada.getRange('F178').setValue('NA');
    Entrada.getRange('F179').setValue('NA');
    Entrada.getRange('F180').setValue('NA');
    //Entrada.getRange('F181').setValue('NA');
    //Entrada.getRange('F182').setValue('NA');
    Entrada.getRange('H137').setValue('NA');
    Entrada.getRange('H138').setValue('NA');
    Entrada.getRange('H139').setValue('NA');
    Entrada.getRange('H140').setValue('NA');
    Entrada.getRange('H141').setValue('NA');
    Entrada.getRange('H142').setValue('NA');
    Entrada.getRange('H143').setValue('NA');
    Entrada.getRange('H144').setValue('NA');
    Entrada.getRange('H145').setValue('NA');
    Entrada.getRange('H146').setValue('NA');
    Entrada.getRange('H147').setValue('NA');
    Entrada.getRange('H148').setValue('NA');
    Entrada.getRange('H149').setValue('NA');
    Entrada.getRange('H150').setValue('NA');
    Entrada.getRange('H151').setValue('NA');
    Entrada.getRange('H152').setValue('NA');
    Entrada.getRange('H153').setValue('NA');
    Entrada.getRange('H154').setValue('NA');
    Entrada.getRange('H155').setValue('NA');
    Entrada.getRange('H156').setValue('NA');
    Entrada.getRange('H157').setValue('NA');
    Entrada.getRange('H158').setValue('NA');
    Entrada.getRange('H159').setValue('NA');
    Entrada.getRange('H160').setValue('NA');
    Entrada.getRange('H161').setValue('NA');
    Entrada.getRange('H162').setValue('NA');
    Entrada.getRange('H163').setValue('NA');
    Entrada.getRange('H164').setValue('NA');
    Entrada.getRange('H165').setValue('NA');
    Entrada.getRange('H166').setValue('NA');
    Entrada.getRange('H167').setValue('NA');
    Entrada.getRange('H168').setValue('NA');
    Entrada.getRange('H169').setValue('NA');
    Entrada.getRange('H170').setValue('NA');
    Entrada.getRange('H171').setValue('NA');
    Entrada.getRange('H172').setValue('NA');
    Entrada.getRange('H173').setValue('NA');
    Entrada.getRange('H174').setValue('NA');
    Entrada.getRange('H175').setValue('NA');
    Entrada.getRange('H176').setValue('NA');
    Entrada.getRange('H177').setValue('NA');
    Entrada.getRange('H178').setValue('NA');
    Entrada.getRange('H179').setValue('NA');
    //Entrada.getRange('H180').setValue('NA');
    //Entrada.getRange('H181').setValue('NA');
    //Entrada.getRange('H182').setValue('NA');
	} else if (acusados == 1){
	//ACUSADO 2     
    Entrada.getRange('F43').setValue('NA');
    Entrada.getRange('F44').setValue('NA');
    Entrada.getRange('F45').setValue('NA');
    Entrada.getRange('F46').setValue('NA');
    Entrada.getRange('F47').setValue('NA');
    Entrada.getRange('F51').setValue('NA');
    Entrada.getRange('F52').setValue('NA');
    Entrada.getRange('F53').setValue('NA');
    Entrada.getRange('F54').setValue('NA');
    Entrada.getRange('F55').setValue('NA');
    Entrada.getRange('F56').setValue('NA');
    Entrada.getRange('F57').setValue('NA');
    Entrada.getRange('F58').setValue('NA');
    Entrada.getRange('F59').setValue('NA');
    Entrada.getRange('F60').setValue('NA');
    Entrada.getRange('F61').setValue('NA');
    //Entrada.getRange('F62').setValue('NA');
    Entrada.getRange('F63').setValue('NA');
    Entrada.getRange('F64').setValue('NA');
    Entrada.getRange('F65').setValue('NA');
    Entrada.getRange('F66').setValue('NA');
    Entrada.getRange('F67').setValue('NA');
    Entrada.getRange('F68').setValue('NA');
    Entrada.getRange('F69').setValue('NA');
    Entrada.getRange('F70').setValue('NA');
    Entrada.getRange('F71').setValue('NA');
    Entrada.getRange('F72').setValue('NA');
    Entrada.getRange('F73').setValue('NA');
    //Entrada.getRange('F74').setValue('NA');
    Entrada.getRange('F75').setValue('NA');
    Entrada.getRange('F76').setValue('NA');
    Entrada.getRange('F77').setValue('NA');
    Entrada.getRange('F78').setValue('NA');
    //Entrada.getRange('F79').setValue('NA');
    Entrada.getRange('F80').setValue('NA');
    Entrada.getRange('F81').setValue('NA');
    Entrada.getRange('F82').setValue('NA');
    Entrada.getRange('F83').setValue('NA');
    Entrada.getRange('F84').setValue('NA');
    Entrada.getRange('F85').setValue('NA');
    Entrada.getRange('F86').setValue('NA');
    //Entrada.getRange('F87').setValue('NA');
    //Entrada.getRange('F88').setValue('NA');
    Entrada.getRange('H43').setValue('NA');
    Entrada.getRange('H44').setValue('NA');
    Entrada.getRange('H45').setValue('NA');
    Entrada.getRange('H46').setValue('NA');
    Entrada.getRange('H47').setValue('NA');
    Entrada.getRange('H48').setValue('NA');
    Entrada.getRange('H49').setValue('NA');
    Entrada.getRange('H50').setValue('NA');
    Entrada.getRange('H51').setValue('NA');
    Entrada.getRange('H52').setValue('NA');
    Entrada.getRange('H53').setValue('NA');
    Entrada.getRange('H54').setValue('NA');
    Entrada.getRange('H55').setValue('NA');
    Entrada.getRange('H56').setValue('NA');
    Entrada.getRange('H57').setValue('NA');
    Entrada.getRange('H58').setValue('NA');
    Entrada.getRange('H59').setValue('NA');
    Entrada.getRange('H60').setValue('NA');
    Entrada.getRange('H61').setValue('NA');
    Entrada.getRange('H62').setValue('NA');
    Entrada.getRange('H63').setValue('NA');
    Entrada.getRange('H64').setValue('NA');
    Entrada.getRange('H65').setValue('NA');
    Entrada.getRange('H66').setValue('NA');
    Entrada.getRange('H67').setValue('NA');
    Entrada.getRange('H68').setValue('NA');
    Entrada.getRange('H69').setValue('NA');
    Entrada.getRange('H70').setValue('NA');
    Entrada.getRange('H71').setValue('NA');
    Entrada.getRange('H72').setValue('NA');
    Entrada.getRange('H73').setValue('NA');
    Entrada.getRange('H74').setValue('NA');
    Entrada.getRange('H75').setValue('NA');
    Entrada.getRange('H76').setValue('NA');
    Entrada.getRange('H77').setValue('NA');
    Entrada.getRange('H78').setValue('NA');
    Entrada.getRange('H79').setValue('NA');
    Entrada.getRange('H80').setValue('NA');
    Entrada.getRange('H81').setValue('NA');
    Entrada.getRange('H82').setValue('NA');
    Entrada.getRange('H83').setValue('NA');
    Entrada.getRange('H84').setValue('NA');
    Entrada.getRange('H85').setValue('NA');
    //Entrada.getRange('H86').setValue('NA');
    //Entrada.getRange('H87').setValue('NA');
    //Entrada.getRange('H88').setValue('NA');
	//ACUSADO 3
    Entrada.getRange('B90').setValue('NA');
    Entrada.getRange('B91').setValue('NA');
    Entrada.getRange('B92').setValue('NA');
    Entrada.getRange('B93').setValue('NA');
    Entrada.getRange('B94').setValue('NA');
    Entrada.getRange('B98').setValue('NA');
    Entrada.getRange('B99').setValue('NA');
    Entrada.getRange('B100').setValue('NA');
    Entrada.getRange('B101').setValue('NA');
    Entrada.getRange('B102').setValue('NA');
    Entrada.getRange('B103').setValue('NA');
    Entrada.getRange('B104').setValue('NA');
    Entrada.getRange('B105').setValue('NA');
    Entrada.getRange('B106').setValue('NA');
    Entrada.getRange('B107').setValue('NA');
    Entrada.getRange('B108').setValue('NA');
    //Entrada.getRange('B109').setValue('NA');
    Entrada.getRange('B110').setValue('NA');
    Entrada.getRange('B111').setValue('NA');
    Entrada.getRange('B112').setValue('NA');
    Entrada.getRange('B113').setValue('NA');
    Entrada.getRange('B114').setValue('NA');
    Entrada.getRange('B115').setValue('NA');
    Entrada.getRange('B116').setValue('NA');
    Entrada.getRange('B117').setValue('NA');
    Entrada.getRange('B118').setValue('NA');
    Entrada.getRange('B119').setValue('NA');
    Entrada.getRange('B120').setValue('NA');
    //Entrada.getRange('B121').setValue('NA');
    Entrada.getRange('B122').setValue('NA');
    Entrada.getRange('B123').setValue('NA');
    Entrada.getRange('B124').setValue('NA');
    Entrada.getRange('B125').setValue('NA');
    //Entrada.getRange('B126').setValue('NA');
    Entrada.getRange('B127').setValue('NA');
    Entrada.getRange('B128').setValue('NA');
    Entrada.getRange('B129').setValue('NA');
    Entrada.getRange('B130').setValue('NA');
    Entrada.getRange('B131').setValue('NA');
    Entrada.getRange('B132').setValue('NA');
    Entrada.getRange('B133').setValue('NA');
    //Entrada.getRange('B134').setValue('NA');
    //Entrada.getRange('B135').setValue('NA');
    Entrada.getRange('D90').setValue('NA');
    Entrada.getRange('D91').setValue('NA');
    Entrada.getRange('D92').setValue('NA');
    Entrada.getRange('D93').setValue('NA');
    Entrada.getRange('D94').setValue('NA');
    Entrada.getRange('D95').setValue('NA');
    Entrada.getRange('D96').setValue('NA');
    Entrada.getRange('D97').setValue('NA');
    Entrada.getRange('D98').setValue('NA');
    Entrada.getRange('D99').setValue('NA');
    Entrada.getRange('D100').setValue('NA');
    Entrada.getRange('D101').setValue('NA');
    Entrada.getRange('D102').setValue('NA');
    Entrada.getRange('D103').setValue('NA');
    Entrada.getRange('D104').setValue('NA');
    Entrada.getRange('D105').setValue('NA');
    Entrada.getRange('D106').setValue('NA');
    Entrada.getRange('D107').setValue('NA');
    Entrada.getRange('D108').setValue('NA');
    Entrada.getRange('D109').setValue('NA');
    Entrada.getRange('D110').setValue('NA');
    Entrada.getRange('D111').setValue('NA');
    Entrada.getRange('D112').setValue('NA');
    Entrada.getRange('D113').setValue('NA');
    Entrada.getRange('D114').setValue('NA');
    Entrada.getRange('D115').setValue('NA');
    Entrada.getRange('D116').setValue('NA');
    Entrada.getRange('D117').setValue('NA');
    Entrada.getRange('D118').setValue('NA');
    Entrada.getRange('D119').setValue('NA');
    Entrada.getRange('D120').setValue('NA');
    Entrada.getRange('D121').setValue('NA');
    Entrada.getRange('D122').setValue('NA');
    Entrada.getRange('D123').setValue('NA');
    Entrada.getRange('D124').setValue('NA');
    Entrada.getRange('D125').setValue('NA');
    Entrada.getRange('D126').setValue('NA');
    Entrada.getRange('D127').setValue('NA');
    Entrada.getRange('D128').setValue('NA');
    Entrada.getRange('D129').setValue('NA');
    Entrada.getRange('D130').setValue('NA');
    Entrada.getRange('D131').setValue('NA');
    Entrada.getRange('D132').setValue('NA');
    //Entrada.getRange('D133').setValue('NA');
    //Entrada.getRange('D134').setValue('NA');
    //Entrada.getRange('D135').setValue('NA');
	//ACUSADO 4
    Entrada.getRange('F90').setValue('NA');
    Entrada.getRange('F91').setValue('NA');
    Entrada.getRange('F92').setValue('NA');
    Entrada.getRange('F93').setValue('NA');
    Entrada.getRange('F94').setValue('NA');
    Entrada.getRange('F98').setValue('NA');
    Entrada.getRange('F99').setValue('NA');
    Entrada.getRange('F100').setValue('NA');
    Entrada.getRange('F101').setValue('NA');
    Entrada.getRange('F102').setValue('NA');
    Entrada.getRange('F103').setValue('NA');
    Entrada.getRange('F104').setValue('NA');
    Entrada.getRange('F105').setValue('NA');
    Entrada.getRange('F106').setValue('NA');
    Entrada.getRange('F107').setValue('NA');
    Entrada.getRange('F108').setValue('NA');
    //Entrada.getRange('F109').setValue('NA');
    Entrada.getRange('F110').setValue('NA');
    Entrada.getRange('F111').setValue('NA');
    Entrada.getRange('F112').setValue('NA');
    Entrada.getRange('F113').setValue('NA');
    Entrada.getRange('F114').setValue('NA');
    Entrada.getRange('F115').setValue('NA');
    Entrada.getRange('F116').setValue('NA');
    Entrada.getRange('F117').setValue('NA');
    Entrada.getRange('F118').setValue('NA');
    Entrada.getRange('F119').setValue('NA');
    Entrada.getRange('F120').setValue('NA');
    //Entrada.getRange('F121').setValue('NA');
    Entrada.getRange('F122').setValue('NA');
    Entrada.getRange('F123').setValue('NA');
    Entrada.getRange('F124').setValue('NA');
    Entrada.getRange('F125').setValue('NA');
    //Entrada.getRange('F126').setValue('NA');
    Entrada.getRange('F127').setValue('NA');
    Entrada.getRange('F128').setValue('NA');
    Entrada.getRange('F129').setValue('NA');
    Entrada.getRange('F130').setValue('NA');
    Entrada.getRange('F131').setValue('NA');
    Entrada.getRange('F132').setValue('NA');
    Entrada.getRange('F133').setValue('NA');
    //Entrada.getRange('F134').setValue('NA');
    //Entrada.getRange('F135').setValue('NA');
    Entrada.getRange('H90').setValue('NA');
    Entrada.getRange('H91').setValue('NA');
    Entrada.getRange('H92').setValue('NA');
    Entrada.getRange('H93').setValue('NA');
    Entrada.getRange('H94').setValue('NA');
    Entrada.getRange('H95').setValue('NA');
    Entrada.getRange('H96').setValue('NA');
    Entrada.getRange('H97').setValue('NA');
    Entrada.getRange('H98').setValue('NA');
    Entrada.getRange('H99').setValue('NA');
    Entrada.getRange('H100').setValue('NA');
    Entrada.getRange('H101').setValue('NA');
    Entrada.getRange('H102').setValue('NA');
    Entrada.getRange('H103').setValue('NA');
    Entrada.getRange('H104').setValue('NA');
    Entrada.getRange('H105').setValue('NA');
    Entrada.getRange('H106').setValue('NA');
    Entrada.getRange('H107').setValue('NA');
    Entrada.getRange('H108').setValue('NA');
    Entrada.getRange('H109').setValue('NA');
    Entrada.getRange('H110').setValue('NA');
    Entrada.getRange('H111').setValue('NA');
    Entrada.getRange('H112').setValue('NA');
    Entrada.getRange('H113').setValue('NA');
    Entrada.getRange('H114').setValue('NA');
    Entrada.getRange('H115').setValue('NA');
    Entrada.getRange('H116').setValue('NA');
    Entrada.getRange('H117').setValue('NA');
    Entrada.getRange('H118').setValue('NA');
    Entrada.getRange('H119').setValue('NA');
    Entrada.getRange('H120').setValue('NA');
    Entrada.getRange('H121').setValue('NA');
    Entrada.getRange('H122').setValue('NA');
    Entrada.getRange('H123').setValue('NA');
    Entrada.getRange('H124').setValue('NA');
    Entrada.getRange('H125').setValue('NA');
    Entrada.getRange('H126').setValue('NA');
    Entrada.getRange('H127').setValue('NA');
    Entrada.getRange('H128').setValue('NA');
    Entrada.getRange('H129').setValue('NA');
    Entrada.getRange('H130').setValue('NA');
    Entrada.getRange('H131').setValue('NA');
    Entrada.getRange('H132').setValue('NA');
    //Entrada.getRange('H133').setValue('NA');
    //Entrada.getRange('H134').setValue('NA');
    //Entrada.getRange('H135').setValue('NA');
	//ACUSADO 5     
    Entrada.getRange('B137').setValue('NA');
    Entrada.getRange('B138').setValue('NA');
    Entrada.getRange('B139').setValue('NA');
    Entrada.getRange('B140').setValue('NA');
    Entrada.getRange('B141').setValue('NA');
    Entrada.getRange('B145').setValue('NA');
    Entrada.getRange('B146').setValue('NA');
    Entrada.getRange('B147').setValue('NA');
    Entrada.getRange('B148').setValue('NA');
    Entrada.getRange('B149').setValue('NA');
    Entrada.getRange('B150').setValue('NA');
    Entrada.getRange('B151').setValue('NA');
    Entrada.getRange('B152').setValue('NA');
    Entrada.getRange('B153').setValue('NA');
    Entrada.getRange('B154').setValue('NA');
    Entrada.getRange('B155').setValue('NA');
    //Entrada.getRange('B156').setValue('NA');
    Entrada.getRange('B157').setValue('NA');
    Entrada.getRange('B158').setValue('NA');
    Entrada.getRange('B159').setValue('NA');
    Entrada.getRange('B160').setValue('NA');
    Entrada.getRange('B161').setValue('NA');
    Entrada.getRange('B162').setValue('NA');
    Entrada.getRange('B163').setValue('NA');
    Entrada.getRange('B164').setValue('NA');
    Entrada.getRange('B165').setValue('NA');
    Entrada.getRange('B166').setValue('NA');
    Entrada.getRange('B167').setValue('NA');
    //Entrada.getRange('B168').setValue('NA');
    Entrada.getRange('B169').setValue('NA');
    Entrada.getRange('B170').setValue('NA');
    Entrada.getRange('B171').setValue('NA');
    Entrada.getRange('B172').setValue('NA');
    //Entrada.getRange('B173').setValue('NA');
    Entrada.getRange('B174').setValue('NA');
    Entrada.getRange('B175').setValue('NA');
    Entrada.getRange('B176').setValue('NA');
    Entrada.getRange('B177').setValue('NA');
    Entrada.getRange('B178').setValue('NA');
    Entrada.getRange('B179').setValue('NA');
    Entrada.getRange('B180').setValue('NA');
    //Entrada.getRange('B181').setValue('NA');
    //Entrada.getRange('B182').setValue('NA');
    Entrada.getRange('D137').setValue('NA');
    Entrada.getRange('D138').setValue('NA');
    Entrada.getRange('D139').setValue('NA');
    Entrada.getRange('D140').setValue('NA');
    Entrada.getRange('D141').setValue('NA');
    Entrada.getRange('D142').setValue('NA');
    Entrada.getRange('D143').setValue('NA');
    Entrada.getRange('D144').setValue('NA');
    Entrada.getRange('D145').setValue('NA');
    Entrada.getRange('D146').setValue('NA');
    Entrada.getRange('D147').setValue('NA');
    Entrada.getRange('D148').setValue('NA');
    Entrada.getRange('D149').setValue('NA');
    Entrada.getRange('D150').setValue('NA');
    Entrada.getRange('D151').setValue('NA');
    Entrada.getRange('D152').setValue('NA');
    Entrada.getRange('D153').setValue('NA');
    Entrada.getRange('D154').setValue('NA');
    Entrada.getRange('D155').setValue('NA');
    Entrada.getRange('D156').setValue('NA');
    Entrada.getRange('D157').setValue('NA');
    Entrada.getRange('D158').setValue('NA');
    Entrada.getRange('D159').setValue('NA');
    Entrada.getRange('D160').setValue('NA');
    Entrada.getRange('D161').setValue('NA');
    Entrada.getRange('D162').setValue('NA');
    Entrada.getRange('D163').setValue('NA');
    Entrada.getRange('D164').setValue('NA');
    Entrada.getRange('D165').setValue('NA');
    Entrada.getRange('D166').setValue('NA');
    Entrada.getRange('D167').setValue('NA');
    Entrada.getRange('D168').setValue('NA');
    Entrada.getRange('D169').setValue('NA');
    Entrada.getRange('D170').setValue('NA');
    Entrada.getRange('D171').setValue('NA');
    Entrada.getRange('D172').setValue('NA');
    Entrada.getRange('D173').setValue('NA');
    Entrada.getRange('D174').setValue('NA');
    Entrada.getRange('D175').setValue('NA');
    Entrada.getRange('D176').setValue('NA');
    Entrada.getRange('D177').setValue('NA');
    Entrada.getRange('D178').setValue('NA');
    Entrada.getRange('D179').setValue('NA');
    //Entrada.getRange('D180').setValue('NA');
    //Entrada.getRange('D181').setValue('NA');
    //Entrada.getRange('D182').setValue('NA');
	//ACUSADO 6     
    Entrada.getRange('F137').setValue('NA');
    Entrada.getRange('F138').setValue('NA');
    Entrada.getRange('F139').setValue('NA');
    Entrada.getRange('F140').setValue('NA');
    Entrada.getRange('F141').setValue('NA');
    Entrada.getRange('F145').setValue('NA');
    Entrada.getRange('F146').setValue('NA');
    Entrada.getRange('F147').setValue('NA');
    Entrada.getRange('F148').setValue('NA');
    Entrada.getRange('F149').setValue('NA');
    Entrada.getRange('F150').setValue('NA');
    Entrada.getRange('F151').setValue('NA');
    Entrada.getRange('F152').setValue('NA');
    Entrada.getRange('F153').setValue('NA');
    Entrada.getRange('F154').setValue('NA');
    Entrada.getRange('F155').setValue('NA');
    //Entrada.getRange('F156').setValue('NA');
    Entrada.getRange('F157').setValue('NA');
    Entrada.getRange('F158').setValue('NA');
    Entrada.getRange('F159').setValue('NA');
    Entrada.getRange('F160').setValue('NA');
    Entrada.getRange('F161').setValue('NA');
    Entrada.getRange('F162').setValue('NA');
    Entrada.getRange('F163').setValue('NA');
    Entrada.getRange('F164').setValue('NA');
    Entrada.getRange('F165').setValue('NA');
    Entrada.getRange('F166').setValue('NA');
    Entrada.getRange('F167').setValue('NA');
    //Entrada.getRange('F168').setValue('NA');
    Entrada.getRange('F169').setValue('NA');
    Entrada.getRange('F170').setValue('NA');
    Entrada.getRange('F171').setValue('NA');
    Entrada.getRange('F172').setValue('NA');
    //Entrada.getRange('F173').setValue('NA');
    Entrada.getRange('F174').setValue('NA');
    Entrada.getRange('F175').setValue('NA');
    Entrada.getRange('F176').setValue('NA');
    Entrada.getRange('F177').setValue('NA');
    Entrada.getRange('F178').setValue('NA');
    Entrada.getRange('F179').setValue('NA');
    Entrada.getRange('F180').setValue('NA');
    //Entrada.getRange('F181').setValue('NA');
    //Entrada.getRange('F182').setValue('NA');
    Entrada.getRange('H137').setValue('NA');
    Entrada.getRange('H138').setValue('NA');
    Entrada.getRange('H139').setValue('NA');
    Entrada.getRange('H140').setValue('NA');
    Entrada.getRange('H141').setValue('NA');
    Entrada.getRange('H142').setValue('NA');
    Entrada.getRange('H143').setValue('NA');
    Entrada.getRange('H144').setValue('NA');
    Entrada.getRange('H145').setValue('NA');
    Entrada.getRange('H146').setValue('NA');
    Entrada.getRange('H147').setValue('NA');
    Entrada.getRange('H148').setValue('NA');
    Entrada.getRange('H149').setValue('NA');
    Entrada.getRange('H150').setValue('NA');
    Entrada.getRange('H151').setValue('NA');
    Entrada.getRange('H152').setValue('NA');
    Entrada.getRange('H153').setValue('NA');
    Entrada.getRange('H154').setValue('NA');
    Entrada.getRange('H155').setValue('NA');
    Entrada.getRange('H156').setValue('NA');
    Entrada.getRange('H157').setValue('NA');
    Entrada.getRange('H158').setValue('NA');
    Entrada.getRange('H159').setValue('NA');
    Entrada.getRange('H160').setValue('NA');
    Entrada.getRange('H161').setValue('NA');
    Entrada.getRange('H162').setValue('NA');
    Entrada.getRange('H163').setValue('NA');
    Entrada.getRange('H164').setValue('NA');
    Entrada.getRange('H165').setValue('NA');
    Entrada.getRange('H166').setValue('NA');
    Entrada.getRange('H167').setValue('NA');
    Entrada.getRange('H168').setValue('NA');
    Entrada.getRange('H169').setValue('NA');
    Entrada.getRange('H170').setValue('NA');
    Entrada.getRange('H171').setValue('NA');
    Entrada.getRange('H172').setValue('NA');
    Entrada.getRange('H173').setValue('NA');
    Entrada.getRange('H174').setValue('NA');
    Entrada.getRange('H175').setValue('NA');
    Entrada.getRange('H176').setValue('NA');
    Entrada.getRange('H177').setValue('NA');
    Entrada.getRange('H178').setValue('NA');
    Entrada.getRange('H179').setValue('NA');
    //Entrada.getRange('H180').setValue('NA');
    //Entrada.getRange('H181').setValue('NA');
    //Entrada.getRange('H182').setValue('NA');
	} else if (acusados == 2){
	//ACUSADO 3
    Entrada.getRange('B90').setValue('NA');
    Entrada.getRange('B91').setValue('NA');
    Entrada.getRange('B92').setValue('NA');
    Entrada.getRange('B93').setValue('NA');
    Entrada.getRange('B94').setValue('NA');
    Entrada.getRange('B98').setValue('NA');
    Entrada.getRange('B99').setValue('NA');
    Entrada.getRange('B100').setValue('NA');
    Entrada.getRange('B101').setValue('NA');
    Entrada.getRange('B102').setValue('NA');
    Entrada.getRange('B103').setValue('NA');
    Entrada.getRange('B104').setValue('NA');
    Entrada.getRange('B105').setValue('NA');
    Entrada.getRange('B106').setValue('NA');
    Entrada.getRange('B107').setValue('NA');
    Entrada.getRange('B108').setValue('NA');
    //Entrada.getRange('B109').setValue('NA');
    Entrada.getRange('B110').setValue('NA');
    Entrada.getRange('B111').setValue('NA');
    Entrada.getRange('B112').setValue('NA');
    Entrada.getRange('B113').setValue('NA');
    Entrada.getRange('B114').setValue('NA');
    Entrada.getRange('B115').setValue('NA');
    Entrada.getRange('B116').setValue('NA');
    Entrada.getRange('B117').setValue('NA');
    Entrada.getRange('B118').setValue('NA');
    Entrada.getRange('B119').setValue('NA');
    Entrada.getRange('B120').setValue('NA');
    //Entrada.getRange('B121').setValue('NA');
    Entrada.getRange('B122').setValue('NA');
    Entrada.getRange('B123').setValue('NA');
    Entrada.getRange('B124').setValue('NA');
    Entrada.getRange('B125').setValue('NA');
    //Entrada.getRange('B126').setValue('NA');
    Entrada.getRange('B127').setValue('NA');
    Entrada.getRange('B128').setValue('NA');
    Entrada.getRange('B129').setValue('NA');
    Entrada.getRange('B130').setValue('NA');
    Entrada.getRange('B131').setValue('NA');
    Entrada.getRange('B132').setValue('NA');
    Entrada.getRange('B133').setValue('NA');
    //Entrada.getRange('B134').setValue('NA');
    //Entrada.getRange('B135').setValue('NA');
    Entrada.getRange('D90').setValue('NA');
    Entrada.getRange('D91').setValue('NA');
    Entrada.getRange('D92').setValue('NA');
    Entrada.getRange('D93').setValue('NA');
    Entrada.getRange('D94').setValue('NA');
    Entrada.getRange('D95').setValue('NA');
    Entrada.getRange('D96').setValue('NA');
    Entrada.getRange('D97').setValue('NA');
    Entrada.getRange('D98').setValue('NA');
    Entrada.getRange('D99').setValue('NA');
    Entrada.getRange('D100').setValue('NA');
    Entrada.getRange('D101').setValue('NA');
    Entrada.getRange('D102').setValue('NA');
    Entrada.getRange('D103').setValue('NA');
    Entrada.getRange('D104').setValue('NA');
    Entrada.getRange('D105').setValue('NA');
    Entrada.getRange('D106').setValue('NA');
    Entrada.getRange('D107').setValue('NA');
    Entrada.getRange('D108').setValue('NA');
    Entrada.getRange('D109').setValue('NA');
    Entrada.getRange('D110').setValue('NA');
    Entrada.getRange('D111').setValue('NA');
    Entrada.getRange('D112').setValue('NA');
    Entrada.getRange('D113').setValue('NA');
    Entrada.getRange('D114').setValue('NA');
    Entrada.getRange('D115').setValue('NA');
    Entrada.getRange('D116').setValue('NA');
    Entrada.getRange('D117').setValue('NA');
    Entrada.getRange('D118').setValue('NA');
    Entrada.getRange('D119').setValue('NA');
    Entrada.getRange('D120').setValue('NA');
    Entrada.getRange('D121').setValue('NA');
    Entrada.getRange('D122').setValue('NA');
    Entrada.getRange('D123').setValue('NA');
    Entrada.getRange('D124').setValue('NA');
    Entrada.getRange('D125').setValue('NA');
    Entrada.getRange('D126').setValue('NA');
    Entrada.getRange('D127').setValue('NA');
    Entrada.getRange('D128').setValue('NA');
    Entrada.getRange('D129').setValue('NA');
    Entrada.getRange('D130').setValue('NA');
    Entrada.getRange('D131').setValue('NA');
    Entrada.getRange('D132').setValue('NA');
    //Entrada.getRange('D133').setValue('NA');
    //Entrada.getRange('D134').setValue('NA');
    //Entrada.getRange('D135').setValue('NA');
	//ACUSADO 4
    Entrada.getRange('F90').setValue('NA');
    Entrada.getRange('F91').setValue('NA');
    Entrada.getRange('F92').setValue('NA');
    Entrada.getRange('F93').setValue('NA');
    Entrada.getRange('F94').setValue('NA');
    Entrada.getRange('F98').setValue('NA');
    Entrada.getRange('F99').setValue('NA');
    Entrada.getRange('F100').setValue('NA');
    Entrada.getRange('F101').setValue('NA');
    Entrada.getRange('F102').setValue('NA');
    Entrada.getRange('F103').setValue('NA');
    Entrada.getRange('F104').setValue('NA');
    Entrada.getRange('F105').setValue('NA');
    Entrada.getRange('F106').setValue('NA');
    Entrada.getRange('F107').setValue('NA');
    Entrada.getRange('F108').setValue('NA');
    //Entrada.getRange('F109').setValue('NA');
    Entrada.getRange('F110').setValue('NA');
    Entrada.getRange('F111').setValue('NA');
    Entrada.getRange('F112').setValue('NA');
    Entrada.getRange('F113').setValue('NA');
    Entrada.getRange('F114').setValue('NA');
    Entrada.getRange('F115').setValue('NA');
    Entrada.getRange('F116').setValue('NA');
    Entrada.getRange('F117').setValue('NA');
    Entrada.getRange('F118').setValue('NA');
    Entrada.getRange('F119').setValue('NA');
    Entrada.getRange('F120').setValue('NA');
    //Entrada.getRange('F121').setValue('NA');
    Entrada.getRange('F122').setValue('NA');
    Entrada.getRange('F123').setValue('NA');
    Entrada.getRange('F124').setValue('NA');
    Entrada.getRange('F125').setValue('NA');
    //Entrada.getRange('F126').setValue('NA');
    Entrada.getRange('F127').setValue('NA');
    Entrada.getRange('F128').setValue('NA');
    Entrada.getRange('F129').setValue('NA');
    Entrada.getRange('F130').setValue('NA');
    Entrada.getRange('F131').setValue('NA');
    Entrada.getRange('F132').setValue('NA');
    Entrada.getRange('F133').setValue('NA');
    //Entrada.getRange('F134').setValue('NA');
    //Entrada.getRange('F135').setValue('NA');
    Entrada.getRange('H90').setValue('NA');
    Entrada.getRange('H91').setValue('NA');
    Entrada.getRange('H92').setValue('NA');
    Entrada.getRange('H93').setValue('NA');
    Entrada.getRange('H94').setValue('NA');
    Entrada.getRange('H95').setValue('NA');
    Entrada.getRange('H96').setValue('NA');
    Entrada.getRange('H97').setValue('NA');
    Entrada.getRange('H98').setValue('NA');
    Entrada.getRange('H99').setValue('NA');
    Entrada.getRange('H100').setValue('NA');
    Entrada.getRange('H101').setValue('NA');
    Entrada.getRange('H102').setValue('NA');
    Entrada.getRange('H103').setValue('NA');
    Entrada.getRange('H104').setValue('NA');
    Entrada.getRange('H105').setValue('NA');
    Entrada.getRange('H106').setValue('NA');
    Entrada.getRange('H107').setValue('NA');
    Entrada.getRange('H108').setValue('NA');
    Entrada.getRange('H109').setValue('NA');
    Entrada.getRange('H110').setValue('NA');
    Entrada.getRange('H111').setValue('NA');
    Entrada.getRange('H112').setValue('NA');
    Entrada.getRange('H113').setValue('NA');
    Entrada.getRange('H114').setValue('NA');
    Entrada.getRange('H115').setValue('NA');
    Entrada.getRange('H116').setValue('NA');
    Entrada.getRange('H117').setValue('NA');
    Entrada.getRange('H118').setValue('NA');
    Entrada.getRange('H119').setValue('NA');
    Entrada.getRange('H120').setValue('NA');
    Entrada.getRange('H121').setValue('NA');
    Entrada.getRange('H122').setValue('NA');
    Entrada.getRange('H123').setValue('NA');
    Entrada.getRange('H124').setValue('NA');
    Entrada.getRange('H125').setValue('NA');
    Entrada.getRange('H126').setValue('NA');
    Entrada.getRange('H127').setValue('NA');
    Entrada.getRange('H128').setValue('NA');
    Entrada.getRange('H129').setValue('NA');
    Entrada.getRange('H130').setValue('NA');
    Entrada.getRange('H131').setValue('NA');
    Entrada.getRange('H132').setValue('NA');
    //Entrada.getRange('H133').setValue('NA');
    //Entrada.getRange('H134').setValue('NA');
    //Entrada.getRange('H135').setValue('NA');
	//ACUSADO 5     
    Entrada.getRange('B137').setValue('NA');
    Entrada.getRange('B138').setValue('NA');
    Entrada.getRange('B139').setValue('NA');
    Entrada.getRange('B140').setValue('NA');
    Entrada.getRange('B141').setValue('NA');
    Entrada.getRange('B145').setValue('NA');
    Entrada.getRange('B146').setValue('NA');
    Entrada.getRange('B147').setValue('NA');
    Entrada.getRange('B148').setValue('NA');
    Entrada.getRange('B149').setValue('NA');
    Entrada.getRange('B150').setValue('NA');
    Entrada.getRange('B151').setValue('NA');
    Entrada.getRange('B152').setValue('NA');
    Entrada.getRange('B153').setValue('NA');
    Entrada.getRange('B154').setValue('NA');
    Entrada.getRange('B155').setValue('NA');
    //Entrada.getRange('B156').setValue('NA');
    Entrada.getRange('B157').setValue('NA');
    Entrada.getRange('B158').setValue('NA');
    Entrada.getRange('B159').setValue('NA');
    Entrada.getRange('B160').setValue('NA');
    Entrada.getRange('B161').setValue('NA');
    Entrada.getRange('B162').setValue('NA');
    Entrada.getRange('B163').setValue('NA');
    Entrada.getRange('B164').setValue('NA');
    Entrada.getRange('B165').setValue('NA');
    Entrada.getRange('B166').setValue('NA');
    Entrada.getRange('B167').setValue('NA');
    //Entrada.getRange('B168').setValue('NA');
    Entrada.getRange('B169').setValue('NA');
    Entrada.getRange('B170').setValue('NA');
    Entrada.getRange('B171').setValue('NA');
    Entrada.getRange('B172').setValue('NA');
    //Entrada.getRange('B173').setValue('NA');
    Entrada.getRange('B174').setValue('NA');
    Entrada.getRange('B175').setValue('NA');
    Entrada.getRange('B176').setValue('NA');
    Entrada.getRange('B177').setValue('NA');
    Entrada.getRange('B178').setValue('NA');
    Entrada.getRange('B179').setValue('NA');
    Entrada.getRange('B180').setValue('NA');
    //Entrada.getRange('B181').setValue('NA');
    //Entrada.getRange('B182').setValue('NA');
    Entrada.getRange('D137').setValue('NA');
    Entrada.getRange('D138').setValue('NA');
    Entrada.getRange('D139').setValue('NA');
    Entrada.getRange('D140').setValue('NA');
    Entrada.getRange('D141').setValue('NA');
    Entrada.getRange('D142').setValue('NA');
    Entrada.getRange('D143').setValue('NA');
    Entrada.getRange('D144').setValue('NA');
    Entrada.getRange('D145').setValue('NA');
    Entrada.getRange('D146').setValue('NA');
    Entrada.getRange('D147').setValue('NA');
    Entrada.getRange('D148').setValue('NA');
    Entrada.getRange('D149').setValue('NA');
    Entrada.getRange('D150').setValue('NA');
    Entrada.getRange('D151').setValue('NA');
    Entrada.getRange('D152').setValue('NA');
    Entrada.getRange('D153').setValue('NA');
    Entrada.getRange('D154').setValue('NA');
    Entrada.getRange('D155').setValue('NA');
    Entrada.getRange('D156').setValue('NA');
    Entrada.getRange('D157').setValue('NA');
    Entrada.getRange('D158').setValue('NA');
    Entrada.getRange('D159').setValue('NA');
    Entrada.getRange('D160').setValue('NA');
    Entrada.getRange('D161').setValue('NA');
    Entrada.getRange('D162').setValue('NA');
    Entrada.getRange('D163').setValue('NA');
    Entrada.getRange('D164').setValue('NA');
    Entrada.getRange('D165').setValue('NA');
    Entrada.getRange('D166').setValue('NA');
    Entrada.getRange('D167').setValue('NA');
    Entrada.getRange('D168').setValue('NA');
    Entrada.getRange('D169').setValue('NA');
    Entrada.getRange('D170').setValue('NA');
    Entrada.getRange('D171').setValue('NA');
    Entrada.getRange('D172').setValue('NA');
    Entrada.getRange('D173').setValue('NA');
    Entrada.getRange('D174').setValue('NA');
    Entrada.getRange('D175').setValue('NA');
    Entrada.getRange('D176').setValue('NA');
    Entrada.getRange('D177').setValue('NA');
    Entrada.getRange('D178').setValue('NA');
    Entrada.getRange('D179').setValue('NA');
    //Entrada.getRange('D180').setValue('NA');
    //Entrada.getRange('D181').setValue('NA');
    //Entrada.getRange('D182').setValue('NA');
	//ACUSADO 6     
    Entrada.getRange('F137').setValue('NA');
    Entrada.getRange('F138').setValue('NA');
    Entrada.getRange('F139').setValue('NA');
    Entrada.getRange('F140').setValue('NA');
    Entrada.getRange('F141').setValue('NA');
    Entrada.getRange('F145').setValue('NA');
    Entrada.getRange('F146').setValue('NA');
    Entrada.getRange('F147').setValue('NA');
    Entrada.getRange('F148').setValue('NA');
    Entrada.getRange('F149').setValue('NA');
    Entrada.getRange('F150').setValue('NA');
    Entrada.getRange('F151').setValue('NA');
    Entrada.getRange('F152').setValue('NA');
    Entrada.getRange('F153').setValue('NA');
    Entrada.getRange('F154').setValue('NA');
    Entrada.getRange('F155').setValue('NA');
    //Entrada.getRange('F156').setValue('NA');
    Entrada.getRange('F157').setValue('NA');
    Entrada.getRange('F158').setValue('NA');
    Entrada.getRange('F159').setValue('NA');
    Entrada.getRange('F160').setValue('NA');
    Entrada.getRange('F161').setValue('NA');
    Entrada.getRange('F162').setValue('NA');
    Entrada.getRange('F163').setValue('NA');
    Entrada.getRange('F164').setValue('NA');
    Entrada.getRange('F165').setValue('NA');
    Entrada.getRange('F166').setValue('NA');
    Entrada.getRange('F167').setValue('NA');
    //Entrada.getRange('F168').setValue('NA');
    Entrada.getRange('F169').setValue('NA');
    Entrada.getRange('F170').setValue('NA');
    Entrada.getRange('F171').setValue('NA');
    Entrada.getRange('F172').setValue('NA');
    //Entrada.getRange('F173').setValue('NA');
    Entrada.getRange('F174').setValue('NA');
    Entrada.getRange('F175').setValue('NA');
    Entrada.getRange('F176').setValue('NA');
    Entrada.getRange('F177').setValue('NA');
    Entrada.getRange('F178').setValue('NA');
    Entrada.getRange('F179').setValue('NA');
    Entrada.getRange('F180').setValue('NA');
    //Entrada.getRange('F181').setValue('NA');
    //Entrada.getRange('F182').setValue('NA');
    Entrada.getRange('H137').setValue('NA');
    Entrada.getRange('H138').setValue('NA');
    Entrada.getRange('H139').setValue('NA');
    Entrada.getRange('H140').setValue('NA');
    Entrada.getRange('H141').setValue('NA');
    Entrada.getRange('H142').setValue('NA');
    Entrada.getRange('H143').setValue('NA');
    Entrada.getRange('H144').setValue('NA');
    Entrada.getRange('H145').setValue('NA');
    Entrada.getRange('H146').setValue('NA');
    Entrada.getRange('H147').setValue('NA');
    Entrada.getRange('H148').setValue('NA');
    Entrada.getRange('H149').setValue('NA');
    Entrada.getRange('H150').setValue('NA');
    Entrada.getRange('H151').setValue('NA');
    Entrada.getRange('H152').setValue('NA');
    Entrada.getRange('H153').setValue('NA');
    Entrada.getRange('H154').setValue('NA');
    Entrada.getRange('H155').setValue('NA');
    Entrada.getRange('H156').setValue('NA');
    Entrada.getRange('H157').setValue('NA');
    Entrada.getRange('H158').setValue('NA');
    Entrada.getRange('H159').setValue('NA');
    Entrada.getRange('H160').setValue('NA');
    Entrada.getRange('H161').setValue('NA');
    Entrada.getRange('H162').setValue('NA');
    Entrada.getRange('H163').setValue('NA');
    Entrada.getRange('H164').setValue('NA');
    Entrada.getRange('H165').setValue('NA');
    Entrada.getRange('H166').setValue('NA');
    Entrada.getRange('H167').setValue('NA');
    Entrada.getRange('H168').setValue('NA');
    Entrada.getRange('H169').setValue('NA');
    Entrada.getRange('H170').setValue('NA');
    Entrada.getRange('H171').setValue('NA');
    Entrada.getRange('H172').setValue('NA');
    Entrada.getRange('H173').setValue('NA');
    Entrada.getRange('H174').setValue('NA');
    Entrada.getRange('H175').setValue('NA');
    Entrada.getRange('H176').setValue('NA');
    Entrada.getRange('H177').setValue('NA');
    Entrada.getRange('H178').setValue('NA');
    Entrada.getRange('H179').setValue('NA');
    //Entrada.getRange('H180').setValue('NA');
    //Entrada.getRange('H181').setValue('NA');
    //Entrada.getRange('H182').setValue('NA');
	} else if (acusados == 3){
	//ACUSADO 4
    Entrada.getRange('F90').setValue('NA');
    Entrada.getRange('F91').setValue('NA');
    Entrada.getRange('F92').setValue('NA');
    Entrada.getRange('F93').setValue('NA');
    Entrada.getRange('F94').setValue('NA');
    Entrada.getRange('F98').setValue('NA');
    Entrada.getRange('F99').setValue('NA');
    Entrada.getRange('F100').setValue('NA');
    Entrada.getRange('F101').setValue('NA');
    Entrada.getRange('F102').setValue('NA');
    Entrada.getRange('F103').setValue('NA');
    Entrada.getRange('F104').setValue('NA');
    Entrada.getRange('F105').setValue('NA');
    Entrada.getRange('F106').setValue('NA');
    Entrada.getRange('F107').setValue('NA');
    Entrada.getRange('F108').setValue('NA');
    //Entrada.getRange('F109').setValue('NA');
    Entrada.getRange('F110').setValue('NA');
    Entrada.getRange('F111').setValue('NA');
    Entrada.getRange('F112').setValue('NA');
    Entrada.getRange('F113').setValue('NA');
    Entrada.getRange('F114').setValue('NA');
    Entrada.getRange('F115').setValue('NA');
    Entrada.getRange('F116').setValue('NA');
    Entrada.getRange('F117').setValue('NA');
    Entrada.getRange('F118').setValue('NA');
    Entrada.getRange('F119').setValue('NA');
    Entrada.getRange('F120').setValue('NA');
    //Entrada.getRange('F121').setValue('NA');
    Entrada.getRange('F122').setValue('NA');
    Entrada.getRange('F123').setValue('NA');
    Entrada.getRange('F124').setValue('NA');
    Entrada.getRange('F125').setValue('NA');
    //Entrada.getRange('F126').setValue('NA');
    Entrada.getRange('F127').setValue('NA');
    Entrada.getRange('F128').setValue('NA');
    Entrada.getRange('F129').setValue('NA');
    Entrada.getRange('F130').setValue('NA');
    Entrada.getRange('F131').setValue('NA');
    Entrada.getRange('F132').setValue('NA');
    Entrada.getRange('F133').setValue('NA');
    //Entrada.getRange('F134').setValue('NA');
    //Entrada.getRange('F135').setValue('NA');
    Entrada.getRange('H90').setValue('NA');
    Entrada.getRange('H91').setValue('NA');
    Entrada.getRange('H92').setValue('NA');
    Entrada.getRange('H93').setValue('NA');
    Entrada.getRange('H94').setValue('NA');
    Entrada.getRange('H95').setValue('NA');
    Entrada.getRange('H96').setValue('NA');
    Entrada.getRange('H97').setValue('NA');
    Entrada.getRange('H98').setValue('NA');
    Entrada.getRange('H99').setValue('NA');
    Entrada.getRange('H100').setValue('NA');
    Entrada.getRange('H101').setValue('NA');
    Entrada.getRange('H102').setValue('NA');
    Entrada.getRange('H103').setValue('NA');
    Entrada.getRange('H104').setValue('NA');
    Entrada.getRange('H105').setValue('NA');
    Entrada.getRange('H106').setValue('NA');
    Entrada.getRange('H107').setValue('NA');
    Entrada.getRange('H108').setValue('NA');
    Entrada.getRange('H109').setValue('NA');
    Entrada.getRange('H110').setValue('NA');
    Entrada.getRange('H111').setValue('NA');
    Entrada.getRange('H112').setValue('NA');
    Entrada.getRange('H113').setValue('NA');
    Entrada.getRange('H114').setValue('NA');
    Entrada.getRange('H115').setValue('NA');
    Entrada.getRange('H116').setValue('NA');
    Entrada.getRange('H117').setValue('NA');
    Entrada.getRange('H118').setValue('NA');
    Entrada.getRange('H119').setValue('NA');
    Entrada.getRange('H120').setValue('NA');
    Entrada.getRange('H121').setValue('NA');
    Entrada.getRange('H122').setValue('NA');
    Entrada.getRange('H123').setValue('NA');
    Entrada.getRange('H124').setValue('NA');
    Entrada.getRange('H125').setValue('NA');
    Entrada.getRange('H126').setValue('NA');
    Entrada.getRange('H127').setValue('NA');
    Entrada.getRange('H128').setValue('NA');
    Entrada.getRange('H129').setValue('NA');
    Entrada.getRange('H130').setValue('NA');
    Entrada.getRange('H131').setValue('NA');
    Entrada.getRange('H132').setValue('NA');
    //Entrada.getRange('H133').setValue('NA');
    //Entrada.getRange('H134').setValue('NA');
    //Entrada.getRange('H135').setValue('NA');
	//ACUSADO 5     
    Entrada.getRange('B137').setValue('NA');
    Entrada.getRange('B138').setValue('NA');
    Entrada.getRange('B139').setValue('NA');
    Entrada.getRange('B140').setValue('NA');
    Entrada.getRange('B141').setValue('NA');
    Entrada.getRange('B145').setValue('NA');
    Entrada.getRange('B146').setValue('NA');
    Entrada.getRange('B147').setValue('NA');
    Entrada.getRange('B148').setValue('NA');
    Entrada.getRange('B149').setValue('NA');
    Entrada.getRange('B150').setValue('NA');
    Entrada.getRange('B151').setValue('NA');
    Entrada.getRange('B152').setValue('NA');
    Entrada.getRange('B153').setValue('NA');
    Entrada.getRange('B154').setValue('NA');
    Entrada.getRange('B155').setValue('NA');
    //Entrada.getRange('B156').setValue('NA');
    Entrada.getRange('B157').setValue('NA');
    Entrada.getRange('B158').setValue('NA');
    Entrada.getRange('B159').setValue('NA');
    Entrada.getRange('B160').setValue('NA');
    Entrada.getRange('B161').setValue('NA');
    Entrada.getRange('B162').setValue('NA');
    Entrada.getRange('B163').setValue('NA');
    Entrada.getRange('B164').setValue('NA');
    Entrada.getRange('B165').setValue('NA');
    Entrada.getRange('B166').setValue('NA');
    Entrada.getRange('B167').setValue('NA');
    //Entrada.getRange('B168').setValue('NA');
    Entrada.getRange('B169').setValue('NA');
    Entrada.getRange('B170').setValue('NA');
    Entrada.getRange('B171').setValue('NA');
    Entrada.getRange('B172').setValue('NA');
    //Entrada.getRange('B173').setValue('NA');
    Entrada.getRange('B174').setValue('NA');
    Entrada.getRange('B175').setValue('NA');
    Entrada.getRange('B176').setValue('NA');
    Entrada.getRange('B177').setValue('NA');
    Entrada.getRange('B178').setValue('NA');
    Entrada.getRange('B179').setValue('NA');
    Entrada.getRange('B180').setValue('NA');
    //Entrada.getRange('B181').setValue('NA');
    //Entrada.getRange('B182').setValue('NA');
    Entrada.getRange('D137').setValue('NA');
    Entrada.getRange('D138').setValue('NA');
    Entrada.getRange('D139').setValue('NA');
    Entrada.getRange('D140').setValue('NA');
    Entrada.getRange('D141').setValue('NA');
    Entrada.getRange('D142').setValue('NA');
    Entrada.getRange('D143').setValue('NA');
    Entrada.getRange('D144').setValue('NA');
    Entrada.getRange('D145').setValue('NA');
    Entrada.getRange('D146').setValue('NA');
    Entrada.getRange('D147').setValue('NA');
    Entrada.getRange('D148').setValue('NA');
    Entrada.getRange('D149').setValue('NA');
    Entrada.getRange('D150').setValue('NA');
    Entrada.getRange('D151').setValue('NA');
    Entrada.getRange('D152').setValue('NA');
    Entrada.getRange('D153').setValue('NA');
    Entrada.getRange('D154').setValue('NA');
    Entrada.getRange('D155').setValue('NA');
    Entrada.getRange('D156').setValue('NA');
    Entrada.getRange('D157').setValue('NA');
    Entrada.getRange('D158').setValue('NA');
    Entrada.getRange('D159').setValue('NA');
    Entrada.getRange('D160').setValue('NA');
    Entrada.getRange('D161').setValue('NA');
    Entrada.getRange('D162').setValue('NA');
    Entrada.getRange('D163').setValue('NA');
    Entrada.getRange('D164').setValue('NA');
    Entrada.getRange('D165').setValue('NA');
    Entrada.getRange('D166').setValue('NA');
    Entrada.getRange('D167').setValue('NA');
    Entrada.getRange('D168').setValue('NA');
    Entrada.getRange('D169').setValue('NA');
    Entrada.getRange('D170').setValue('NA');
    Entrada.getRange('D171').setValue('NA');
    Entrada.getRange('D172').setValue('NA');
    Entrada.getRange('D173').setValue('NA');
    Entrada.getRange('D174').setValue('NA');
    Entrada.getRange('D175').setValue('NA');
    Entrada.getRange('D176').setValue('NA');
    Entrada.getRange('D177').setValue('NA');
    Entrada.getRange('D178').setValue('NA');
    Entrada.getRange('D179').setValue('NA');
    //Entrada.getRange('D180').setValue('NA');
    //Entrada.getRange('D181').setValue('NA');
    //Entrada.getRange('D182').setValue('NA');
	//ACUSADO 6     
    Entrada.getRange('F137').setValue('NA');
    Entrada.getRange('F138').setValue('NA');
    Entrada.getRange('F139').setValue('NA');
    Entrada.getRange('F140').setValue('NA');
    Entrada.getRange('F141').setValue('NA');
    Entrada.getRange('F145').setValue('NA');
    Entrada.getRange('F146').setValue('NA');
    Entrada.getRange('F147').setValue('NA');
    Entrada.getRange('F148').setValue('NA');
    Entrada.getRange('F149').setValue('NA');
    Entrada.getRange('F150').setValue('NA');
    Entrada.getRange('F151').setValue('NA');
    Entrada.getRange('F152').setValue('NA');
    Entrada.getRange('F153').setValue('NA');
    Entrada.getRange('F154').setValue('NA');
    Entrada.getRange('F155').setValue('NA');
    //Entrada.getRange('F156').setValue('NA');
    Entrada.getRange('F157').setValue('NA');
    Entrada.getRange('F158').setValue('NA');
    Entrada.getRange('F159').setValue('NA');
    Entrada.getRange('F160').setValue('NA');
    Entrada.getRange('F161').setValue('NA');
    Entrada.getRange('F162').setValue('NA');
    Entrada.getRange('F163').setValue('NA');
    Entrada.getRange('F164').setValue('NA');
    Entrada.getRange('F165').setValue('NA');
    Entrada.getRange('F166').setValue('NA');
    Entrada.getRange('F167').setValue('NA');
    //Entrada.getRange('F168').setValue('NA');
    Entrada.getRange('F169').setValue('NA');
    Entrada.getRange('F170').setValue('NA');
    Entrada.getRange('F171').setValue('NA');
    Entrada.getRange('F172').setValue('NA');
    //Entrada.getRange('F173').setValue('NA');
    Entrada.getRange('F174').setValue('NA');
    Entrada.getRange('F175').setValue('NA');
    Entrada.getRange('F176').setValue('NA');
    Entrada.getRange('F177').setValue('NA');
    Entrada.getRange('F178').setValue('NA');
    Entrada.getRange('F179').setValue('NA');
    Entrada.getRange('F180').setValue('NA');
    //Entrada.getRange('F181').setValue('NA');
    //Entrada.getRange('F182').setValue('NA');
    Entrada.getRange('H137').setValue('NA');
    Entrada.getRange('H138').setValue('NA');
    Entrada.getRange('H139').setValue('NA');
    Entrada.getRange('H140').setValue('NA');
    Entrada.getRange('H141').setValue('NA');
    Entrada.getRange('H142').setValue('NA');
    Entrada.getRange('H143').setValue('NA');
    Entrada.getRange('H144').setValue('NA');
    Entrada.getRange('H145').setValue('NA');
    Entrada.getRange('H146').setValue('NA');
    Entrada.getRange('H147').setValue('NA');
    Entrada.getRange('H148').setValue('NA');
    Entrada.getRange('H149').setValue('NA');
    Entrada.getRange('H150').setValue('NA');
    Entrada.getRange('H151').setValue('NA');
    Entrada.getRange('H152').setValue('NA');
    Entrada.getRange('H153').setValue('NA');
    Entrada.getRange('H154').setValue('NA');
    Entrada.getRange('H155').setValue('NA');
    Entrada.getRange('H156').setValue('NA');
    Entrada.getRange('H157').setValue('NA');
    Entrada.getRange('H158').setValue('NA');
    Entrada.getRange('H159').setValue('NA');
    Entrada.getRange('H160').setValue('NA');
    Entrada.getRange('H161').setValue('NA');
    Entrada.getRange('H162').setValue('NA');
    Entrada.getRange('H163').setValue('NA');
    Entrada.getRange('H164').setValue('NA');
    Entrada.getRange('H165').setValue('NA');
    Entrada.getRange('H166').setValue('NA');
    Entrada.getRange('H167').setValue('NA');
    Entrada.getRange('H168').setValue('NA');
    Entrada.getRange('H169').setValue('NA');
    Entrada.getRange('H170').setValue('NA');
    Entrada.getRange('H171').setValue('NA');
    Entrada.getRange('H172').setValue('NA');
    Entrada.getRange('H173').setValue('NA');
    Entrada.getRange('H174').setValue('NA');
    Entrada.getRange('H175').setValue('NA');
    Entrada.getRange('H176').setValue('NA');
    Entrada.getRange('H177').setValue('NA');
    Entrada.getRange('H178').setValue('NA');
    Entrada.getRange('H179').setValue('NA');
    //Entrada.getRange('H180').setValue('NA');
    //Entrada.getRange('H181').setValue('NA');
    //Entrada.getRange('H182').setValue('NA');
	} else if (acusados == 4){
	//ACUSADO 5     
    Entrada.getRange('B137').setValue('NA');
    Entrada.getRange('B138').setValue('NA');
    Entrada.getRange('B139').setValue('NA');
    Entrada.getRange('B140').setValue('NA');
    Entrada.getRange('B141').setValue('NA');
    Entrada.getRange('B145').setValue('NA');
    Entrada.getRange('B146').setValue('NA');
    Entrada.getRange('B147').setValue('NA');
    Entrada.getRange('B148').setValue('NA');
    Entrada.getRange('B149').setValue('NA');
    Entrada.getRange('B150').setValue('NA');
    Entrada.getRange('B151').setValue('NA');
    Entrada.getRange('B152').setValue('NA');
    Entrada.getRange('B153').setValue('NA');
    Entrada.getRange('B154').setValue('NA');
    Entrada.getRange('B155').setValue('NA');
    //Entrada.getRange('B156').setValue('NA');
    Entrada.getRange('B157').setValue('NA');
    Entrada.getRange('B158').setValue('NA');
    Entrada.getRange('B159').setValue('NA');
    Entrada.getRange('B160').setValue('NA');
    Entrada.getRange('B161').setValue('NA');
    Entrada.getRange('B162').setValue('NA');
    Entrada.getRange('B163').setValue('NA');
    Entrada.getRange('B164').setValue('NA');
    Entrada.getRange('B165').setValue('NA');
    Entrada.getRange('B166').setValue('NA');
    Entrada.getRange('B167').setValue('NA');
    //Entrada.getRange('B168').setValue('NA');
    Entrada.getRange('B169').setValue('NA');
    Entrada.getRange('B170').setValue('NA');
    Entrada.getRange('B171').setValue('NA');
    Entrada.getRange('B172').setValue('NA');
    //Entrada.getRange('B173').setValue('NA');
    Entrada.getRange('B174').setValue('NA');
    Entrada.getRange('B175').setValue('NA');
    Entrada.getRange('B176').setValue('NA');
    Entrada.getRange('B177').setValue('NA');
    Entrada.getRange('B178').setValue('NA');
    Entrada.getRange('B179').setValue('NA');
    Entrada.getRange('B180').setValue('NA');
    //Entrada.getRange('B181').setValue('NA');
    //Entrada.getRange('B182').setValue('NA');
    Entrada.getRange('D137').setValue('NA');
    Entrada.getRange('D138').setValue('NA');
    Entrada.getRange('D139').setValue('NA');
    Entrada.getRange('D140').setValue('NA');
    Entrada.getRange('D141').setValue('NA');
    Entrada.getRange('D142').setValue('NA');
    Entrada.getRange('D143').setValue('NA');
    Entrada.getRange('D144').setValue('NA');
    Entrada.getRange('D145').setValue('NA');
    Entrada.getRange('D146').setValue('NA');
    Entrada.getRange('D147').setValue('NA');
    Entrada.getRange('D148').setValue('NA');
    Entrada.getRange('D149').setValue('NA');
    Entrada.getRange('D150').setValue('NA');
    Entrada.getRange('D151').setValue('NA');
    Entrada.getRange('D152').setValue('NA');
    Entrada.getRange('D153').setValue('NA');
    Entrada.getRange('D154').setValue('NA');
    Entrada.getRange('D155').setValue('NA');
    Entrada.getRange('D156').setValue('NA');
    Entrada.getRange('D157').setValue('NA');
    Entrada.getRange('D158').setValue('NA');
    Entrada.getRange('D159').setValue('NA');
    Entrada.getRange('D160').setValue('NA');
    Entrada.getRange('D161').setValue('NA');
    Entrada.getRange('D162').setValue('NA');
    Entrada.getRange('D163').setValue('NA');
    Entrada.getRange('D164').setValue('NA');
    Entrada.getRange('D165').setValue('NA');
    Entrada.getRange('D166').setValue('NA');
    Entrada.getRange('D167').setValue('NA');
    Entrada.getRange('D168').setValue('NA');
    Entrada.getRange('D169').setValue('NA');
    Entrada.getRange('D170').setValue('NA');
    Entrada.getRange('D171').setValue('NA');
    Entrada.getRange('D172').setValue('NA');
    Entrada.getRange('D173').setValue('NA');
    Entrada.getRange('D174').setValue('NA');
    Entrada.getRange('D175').setValue('NA');
    Entrada.getRange('D176').setValue('NA');
    Entrada.getRange('D177').setValue('NA');
    Entrada.getRange('D178').setValue('NA');
    Entrada.getRange('D179').setValue('NA');
    //Entrada.getRange('D180').setValue('NA');
    //Entrada.getRange('D181').setValue('NA');
    //Entrada.getRange('D182').setValue('NA');
	//ACUSADO 6     
    Entrada.getRange('F137').setValue('NA');
    Entrada.getRange('F138').setValue('NA');
    Entrada.getRange('F139').setValue('NA');
    Entrada.getRange('F140').setValue('NA');
    Entrada.getRange('F141').setValue('NA');
    Entrada.getRange('F145').setValue('NA');
    Entrada.getRange('F146').setValue('NA');
    Entrada.getRange('F147').setValue('NA');
    Entrada.getRange('F148').setValue('NA');
    Entrada.getRange('F149').setValue('NA');
    Entrada.getRange('F150').setValue('NA');
    Entrada.getRange('F151').setValue('NA');
    Entrada.getRange('F152').setValue('NA');
    Entrada.getRange('F153').setValue('NA');
    Entrada.getRange('F154').setValue('NA');
    Entrada.getRange('F155').setValue('NA');
    //Entrada.getRange('F156').setValue('NA');
    Entrada.getRange('F157').setValue('NA');
    Entrada.getRange('F158').setValue('NA');
    Entrada.getRange('F159').setValue('NA');
    Entrada.getRange('F160').setValue('NA');
    Entrada.getRange('F161').setValue('NA');
    Entrada.getRange('F162').setValue('NA');
    Entrada.getRange('F163').setValue('NA');
    Entrada.getRange('F164').setValue('NA');
    Entrada.getRange('F165').setValue('NA');
    Entrada.getRange('F166').setValue('NA');
    Entrada.getRange('F167').setValue('NA');
    //Entrada.getRange('F168').setValue('NA');
    Entrada.getRange('F169').setValue('NA');
    Entrada.getRange('F170').setValue('NA');
    Entrada.getRange('F171').setValue('NA');
    Entrada.getRange('F172').setValue('NA');
    //Entrada.getRange('F173').setValue('NA');
    Entrada.getRange('F174').setValue('NA');
    Entrada.getRange('F175').setValue('NA');
    Entrada.getRange('F176').setValue('NA');
    Entrada.getRange('F177').setValue('NA');
    Entrada.getRange('F178').setValue('NA');
    Entrada.getRange('F179').setValue('NA');
    Entrada.getRange('F180').setValue('NA');
    //Entrada.getRange('F181').setValue('NA');
    //Entrada.getRange('F182').setValue('NA');
    Entrada.getRange('H137').setValue('NA');
    Entrada.getRange('H138').setValue('NA');
    Entrada.getRange('H139').setValue('NA');
    Entrada.getRange('H140').setValue('NA');
    Entrada.getRange('H141').setValue('NA');
    Entrada.getRange('H142').setValue('NA');
    Entrada.getRange('H143').setValue('NA');
    Entrada.getRange('H144').setValue('NA');
    Entrada.getRange('H145').setValue('NA');
    Entrada.getRange('H146').setValue('NA');
    Entrada.getRange('H147').setValue('NA');
    Entrada.getRange('H148').setValue('NA');
    Entrada.getRange('H149').setValue('NA');
    Entrada.getRange('H150').setValue('NA');
    Entrada.getRange('H151').setValue('NA');
    Entrada.getRange('H152').setValue('NA');
    Entrada.getRange('H153').setValue('NA');
    Entrada.getRange('H154').setValue('NA');
    Entrada.getRange('H155').setValue('NA');
    Entrada.getRange('H156').setValue('NA');
    Entrada.getRange('H157').setValue('NA');
    Entrada.getRange('H158').setValue('NA');
    Entrada.getRange('H159').setValue('NA');
    Entrada.getRange('H160').setValue('NA');
    Entrada.getRange('H161').setValue('NA');
    Entrada.getRange('H162').setValue('NA');
    Entrada.getRange('H163').setValue('NA');
    Entrada.getRange('H164').setValue('NA');
    Entrada.getRange('H165').setValue('NA');
    Entrada.getRange('H166').setValue('NA');
    Entrada.getRange('H167').setValue('NA');
    Entrada.getRange('H168').setValue('NA');
    Entrada.getRange('H169').setValue('NA');
    Entrada.getRange('H170').setValue('NA');
    Entrada.getRange('H171').setValue('NA');
    Entrada.getRange('H172').setValue('NA');
    Entrada.getRange('H173').setValue('NA');
    Entrada.getRange('H174').setValue('NA');
    Entrada.getRange('H175').setValue('NA');
    Entrada.getRange('H176').setValue('NA');
    Entrada.getRange('H177').setValue('NA');
    Entrada.getRange('H178').setValue('NA');
    Entrada.getRange('H179').setValue('NA');
    //Entrada.getRange('H180').setValue('NA');
    //Entrada.getRange('H181').setValue('NA');
    //Entrada.getRange('H182').setValue('NA');
	} else if (acusados == 5){
	//ACUSADO 6     
    Entrada.getRange('F137').setValue('NA');
    Entrada.getRange('F138').setValue('NA');
    Entrada.getRange('F139').setValue('NA');
    Entrada.getRange('F140').setValue('NA');
    Entrada.getRange('F141').setValue('NA');
    Entrada.getRange('F145').setValue('NA');
    Entrada.getRange('F146').setValue('NA');
    Entrada.getRange('F147').setValue('NA');
    Entrada.getRange('F148').setValue('NA');
    Entrada.getRange('F149').setValue('NA');
    Entrada.getRange('F150').setValue('NA');
    Entrada.getRange('F151').setValue('NA');
    Entrada.getRange('F152').setValue('NA');
    Entrada.getRange('F153').setValue('NA');
    Entrada.getRange('F154').setValue('NA');
    Entrada.getRange('F155').setValue('NA');
    //Entrada.getRange('F156').setValue('NA');
    Entrada.getRange('F157').setValue('NA');
    Entrada.getRange('F158').setValue('NA');
    Entrada.getRange('F159').setValue('NA');
    Entrada.getRange('F160').setValue('NA');
    Entrada.getRange('F161').setValue('NA');
    Entrada.getRange('F162').setValue('NA');
    Entrada.getRange('F163').setValue('NA');
    Entrada.getRange('F164').setValue('NA');
    Entrada.getRange('F165').setValue('NA');
    Entrada.getRange('F166').setValue('NA');
    Entrada.getRange('F167').setValue('NA');
    //Entrada.getRange('F168').setValue('NA');
    Entrada.getRange('F169').setValue('NA');
    Entrada.getRange('F170').setValue('NA');
    Entrada.getRange('F171').setValue('NA');
    Entrada.getRange('F172').setValue('NA');
    //Entrada.getRange('F173').setValue('NA');
    Entrada.getRange('F174').setValue('NA');
    Entrada.getRange('F175').setValue('NA');
    Entrada.getRange('F176').setValue('NA');
    Entrada.getRange('F177').setValue('NA');
    Entrada.getRange('F178').setValue('NA');
    Entrada.getRange('F179').setValue('NA');
    Entrada.getRange('F180').setValue('NA');
    //Entrada.getRange('F181').setValue('NA');
    //Entrada.getRange('F182').setValue('NA');
    Entrada.getRange('H137').setValue('NA');
    Entrada.getRange('H138').setValue('NA');
    Entrada.getRange('H139').setValue('NA');
    Entrada.getRange('H140').setValue('NA');
    Entrada.getRange('H141').setValue('NA');
    Entrada.getRange('H142').setValue('NA');
    Entrada.getRange('H143').setValue('NA');
    Entrada.getRange('H144').setValue('NA');
    Entrada.getRange('H145').setValue('NA');
    Entrada.getRange('H146').setValue('NA');
    Entrada.getRange('H147').setValue('NA');
    Entrada.getRange('H148').setValue('NA');
    Entrada.getRange('H149').setValue('NA');
    Entrada.getRange('H150').setValue('NA');
    Entrada.getRange('H151').setValue('NA');
    Entrada.getRange('H152').setValue('NA');
    Entrada.getRange('H153').setValue('NA');
    Entrada.getRange('H154').setValue('NA');
    Entrada.getRange('H155').setValue('NA');
    Entrada.getRange('H156').setValue('NA');
    Entrada.getRange('H157').setValue('NA');
    Entrada.getRange('H158').setValue('NA');
    Entrada.getRange('H159').setValue('NA');
    Entrada.getRange('H160').setValue('NA');
    Entrada.getRange('H161').setValue('NA');
    Entrada.getRange('H162').setValue('NA');
    Entrada.getRange('H163').setValue('NA');
    Entrada.getRange('H164').setValue('NA');
    Entrada.getRange('H165').setValue('NA');
    Entrada.getRange('H166').setValue('NA');
    Entrada.getRange('H167').setValue('NA');
    Entrada.getRange('H168').setValue('NA');
    Entrada.getRange('H169').setValue('NA');
    Entrada.getRange('H170').setValue('NA');
    Entrada.getRange('H171').setValue('NA');
    Entrada.getRange('H172').setValue('NA');
    Entrada.getRange('H173').setValue('NA');
    Entrada.getRange('H174').setValue('NA');
    Entrada.getRange('H175').setValue('NA');
    Entrada.getRange('H176').setValue('NA');
    Entrada.getRange('H177').setValue('NA');
    Entrada.getRange('H178').setValue('NA');
    Entrada.getRange('H179').setValue('NA');
    //Entrada.getRange('H180').setValue('NA');
    //Entrada.getRange('H181').setValue('NA');
    //Entrada.getRange('H182').setValue('NA');
	} else {
		Browser.msgBox("Há 6 acusados nessa ocorrência!") 
	}
}



Array.prototype.findIndex = function(SearchNC){
  
  if(SearchNC == "") return false;
  
   for(var i = 0; i <this.length; i++)
     if(this[i]==SearchNC) return i;
    
   return -1;
  
};

function Verificar(){
  
  var Entrada = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Variaveis');
  var Celulas = Entrada.getRange(2,24).getValue();
  
  if(Celulas <= 0){
    return true
  }else{
    return false
  }
    
}

function UnCheck() {
  var Entrada = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Entrada');
  Entrada.getActiveCell();
  var intervalo = Entrada.getRangeList([]);
  
  intervalo.setValue("0");

}

function LimparEntrada(){
  
  var Entrada = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Entrada');
  
  Entrada.getActiveCell();
  
  Entrada.getRangeList(['B2:B20', 'B22:B26', 'B30:B37', 'B39:B43', 'B47:B55', 'D3', 'D6', 'D8', 'D19:D20', 'D22:D37', 'D39:D58', 'F2:F19', 'F22', 'F24:F28', 'F36', 'F39:F43', 'F47:F55', 'H2:H4', 'H19', 'H21', 'H39:H58']).activate();
  Entrada.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  
  Entrada.getRange('B2').activate();
  

}

function Search(){
  
  var Entrada = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Entrada');
  var Pesquisa = Entrada.getRange('B2').getValue(); // celula com a variavel de busca - NC
  var SheetBanco = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Banco');
  SheetBanco.getActiveCell();
  
  var LocalPesquisa = SheetBanco.getRange(2, 2, SheetBanco.getLastRow()).getValues(); //linha/coluna da variavel NC
  var Resultado = LocalPesquisa.Pesquisa(Pesquisa);
  var Linha = Resultado + 2;
  
  if (Resultado != -1){
    
    var	NC = SheetBanco.getRange(Linha, 2).getValue();
    var	NUM_OCOR = SheetBanco.getRange(Linha,4).getValue();
    var	NUM_FATO = 	SheetBanco.getRange(Linha,3).getValue();
    var	N_FATO = SheetBanco.getRange(Linha,5).getValue();
    var	NUM_ORG_REGIST = SheetBanco.getRange(Linha,6).getValue();
    var	N_ORG_REGIST = SheetBanco.getRange(Linha,7).getValue();
    var	ORIGEM_COM = SheetBanco.getRange(Linha,8).getValue();
    var	DATA_COM = SheetBanco.getRange(Linha,9).getValue();
    var	HORA_COM = SheetBanco.getRange(Linha,10).getValue();
    var	ANO_REGIST = SheetBanco.getRange(Linha,11).getValue();
    var	FLAGRANTE = SheetBanco.getRange(Linha,12).getValue();
    var	TENTATIVA = SheetBanco.getRange(Linha,13).getValue();
    var	COMUM =	SheetBanco.getRange(Linha,14).getValue();
    var	COND = SheetBanco.getRange(Linha,15).getValue();
    var	LAUDO =	SheetBanco.getRange(Linha,16).getValue();
    var	NUM_VIT = SheetBanco.getRange(Linha,17).getValue();
    var	NUM_VIT_MORTAS = SheetBanco.getRange(Linha,18).getValue();
    var	NUMTEST = SheetBanco.getRange(Linha,19).getValue();
    var	TIPO_TEST =	SheetBanco.getRange(Linha,20).getValue();
    var	DATA_FATO =	SheetBanco.getRange(Linha,22).getValue();
    var	H_FATO = SheetBanco.getRange(Linha,27).getValue();
    var	MUN = SheetBanco.getRange(Linha,29).getValue();
    var	A_MUN =	SheetBanco.getRange(Linha,40).getValue();
    var	LOG = SheetBanco.getRange(Linha,41).getValue();
    var	NUM = SheetBanco.getRange(Linha,42).getValue();
    var	COMP = SheetBanco.getRange(Linha,43).getValue();
    var	CEP	= SheetBanco.getRange(Linha,44).getValue();
    var	BAIRRO = SheetBanco.getRange(Linha,45).getValue();
    var	LAT	= SheetBanco.getRange(Linha,46).getValue();
    var	LNG = SheetBanco.getRange(Linha,47).getValue();
    var	P_REF = SheetBanco.getRange(Linha,48).getValue();
    var	TIPO_LOCAL = SheetBanco.getRange(Linha,49).getValue();
    var	M_UTIL_AUT = SheetBanco.getRange(Linha,50).getValue();
    var	RECUR = SheetBanco.getRange(Linha,51).getValue();
    var	NUM_AGRE = SheetBanco.getRange(Linha,52).getValue();
    var	M_AUTOR = SheetBanco.getRange(Linha,53).getValue();
    var	OBJ = SheetBanco.getRange(Linha,54).getValue();
    var	ADTNT = SheetBanco.getRange(Linha,55).getValue();
    var	MOT = SheetBanco.getRange(Linha,56).getValue();
    var	R_VIT_AUT = SheetBanco.getRange(Linha,57).getValue();
    var	MID	= SheetBanco.getRange(Linha,58).getValue();
    var	HIST_1 = SheetBanco.getRange(Linha,59).getValue();
    var	HIST_2 = SheetBanco.getRange(Linha,60).getValue();
    var	N_VIT = SheetBanco.getRange(Linha,63).getValue();
    var	RG_VIT = SheetBanco.getRange(Linha,64).getValue();
    var	CPF_VIT = SheetBanco.getRange(Linha,65).getValue();
    var	SX_VIT = SheetBanco.getRange(Linha,66).getValue();
    var	DN_VIT = SheetBanco.getRange(Linha,67).getValue();
    var	ORIEN_SX_VIT = SheetBanco.getRange(Linha,71).getValue();
    var	COR_VIT = SheetBanco.getRange(Linha,72).getValue();
    var	N_PAIVIT = SheetBanco.getRange(Linha,73).getValue();
    var	N_MAEVIT = SheetBanco.getRange(Linha,74).getValue();
    var	EC_VIT = SheetBanco.getRange(Linha,75).getValue();
    var	UE_VIT = SheetBanco.getRange(Linha,76).getValue();
    var	N_FIL_VIT = SheetBanco.getRange(Linha,77).getValue();
    var	NAT_VIT = SheetBanco.getRange(Linha,78).getValue();
    var	NAC_VIT = SheetBanco.getRange(Linha,79).getValue();
    var	COND_FIS_VIT =	SheetBanco.getRange(Linha,80).getValue();
    var	ESC_VIT = SheetBanco.getRange(Linha,81).getValue();
    var	ESC_MAEVIT = SheetBanco.getRange(Linha,82).getValue();
    var	ESC_PAIVIT = SheetBanco.getRange(Linha,83).getValue();
    var	PROF_VIT = SheetBanco.getRange(Linha,85).getValue();
    var	END_PROFVIT = SheetBanco.getRange(Linha,86).getValue();
    var	END_RESVIT = SheetBanco.getRange(Linha,87).getValue();
    var	MUN_VIT = SheetBanco.getRange(Linha,88).getValue();
    var	BAIRRO_VIT = SheetBanco.getRange(Linha,89).getValue();
    var	ANTEPOL_VIT = SheetBanco.getRange(Linha,90).getValue();
    var	ANTECRIM_VIT = SheetBanco.getRange(Linha,91).getValue();
    var	ENVTRAF_VIT = SheetBanco.getRange(Linha,92).getValue();
    var	MP_VIT = SheetBanco.getRange(Linha,93).getValue();
    var	MI_VIT = SheetBanco.getRange(Linha,94).getValue();
    var	FOR_VIT = SheetBanco.getRange(Linha,95).getValue();
    var	N_ACU_1 = SheetBanco.getRange(Linha,96).getValue();
    var	RG_ACU_1 = SheetBanco.getRange(Linha,97).getValue();
    var	CPF_ACU_1 = SheetBanco.getRange(Linha,98).getValue();
    var	SX_ACU_1 = SheetBanco.getRange(Linha,99).getValue();
    var	DN_ACU_1 = 	SheetBanco.getRange(Linha,100).getValue();
    var	ORIEN_SX_ACU_1 = SheetBanco.getRange(Linha,104).getValue();
    var	COR_ACU_1 = SheetBanco.getRange(Linha,105).getValue();
    var	N_PAI_ACU_1 = SheetBanco.getRange(Linha,106).getValue();
    var	N_MAE_ACU_1 = SheetBanco.getRange(Linha,107).getValue();
    var	EC_ACU_1 = SheetBanco.getRange(Linha,108).getValue();
    var	UE_ACU_1 = SheetBanco.getRange(Linha,109).getValue();
    var	N_FIL_ACU_1 = SheetBanco.getRange(Linha,110).getValue();
    var	NAT_ACU_1 = SheetBanco.getRange(Linha,111).getValue();
    var	NAC_ACU_1 = SheetBanco.getRange(Linha,112).getValue();
    var	COND_FIS_ACU_1 = SheetBanco.getRange(Linha,113).getValue();
    var	ESC_ACU_1 = SheetBanco.getRange(Linha,114).getValue();
    var	ESC_MAE_ACU_1 = SheetBanco.getRange(Linha,115).getValue();
    var	ESC_PAI_ACU_1 = SheetBanco.getRange(Linha,116).getValue();
    var	PROF_ACU_1 = SheetBanco.getRange(Linha,117).getValue();
    var	END_PROF_ACU_1 =SheetBanco.getRange(Linha,118).getValue();
    var	END_RES_ACU_1 = SheetBanco.getRange(Linha,119).getValue();
    var	MUN_ACU_1 = SheetBanco.getRange(Linha,120).getValue();
    var	BAIRRO_ACU_1 = SheetBanco.getRange(Linha,121).getValue();
    var	ANTEPOL_ACU_1 = SheetBanco.getRange(Linha,122).getValue();
    var	ANTECRIM_ACU_1 =SheetBanco.getRange(Linha,123).getValue();
    var	ENVTRAF_ACU_1 =SheetBanco.getRange(Linha,124).getValue();
    var	MP_ACU_1 =SheetBanco.getRange(Linha,125).getValue();
    var	MI_ACU_1 =SheetBanco.getRange(Linha,126).getValue();
    var	FOR_ACU_1 = SheetBanco.getRange(Linha,127).getValue();
    var	N_ACU_2	= SheetBanco.getRange(Linha,128).getValue();
    var	RG_ACU_2 = SheetBanco.getRange(Linha,129).getValue();
    var	CPF_ACU_2 = SheetBanco.getRange(Linha,130).getValue();
    var	SX_ACU_2 = SheetBanco.getRange(Linha,131).getValue();
    var	DN_ACU_2 = SheetBanco.getRange(Linha,132).getValue();
    var	ORIEN_SX_ACU_2 = SheetBanco.getRange(Linha,136).getValue();
    var	COR_ACU_2 = SheetBanco.getRange(Linha,137).getValue();
    var	N_PAI_ACU_2 =SheetBanco.getRange(Linha,138).getValue();
    var	N_MAE_ACU_2 =SheetBanco.getRange(Linha,139).getValue();
    var	EC_ACU_2 = SheetBanco.getRange(Linha,140).getValue();
    var	UE_ACU_2 = SheetBanco.getRange(Linha,141).getValue();
    var	N_FIL_ACU_2 =SheetBanco.getRange(Linha,142).getValue();
    var	NAT_ACU_2 = SheetBanco.getRange(Linha,143).getValue();
    var	NAC_ACU_2 = SheetBanco.getRange(Linha,144).getValue();
    var	COND_FIS_ACU_2 =SheetBanco.getRange(Linha,145).getValue();
    var	ESC_ACU_2 = SheetBanco.getRange(Linha,146).getValue();
    var	ESC_MAE_ACU_2 = SheetBanco.getRange(Linha,147).getValue();
    var	ESC_PAI_ACU_2 = SheetBanco.getRange(Linha,148).getValue();
    var	PROF_ACU_2 = SheetBanco.getRange(Linha,149).getValue();
    var	END_PROF_ACU_2 = SheetBanco.getRange(Linha,150).getValue();
    var	END_RES_ACU_2 = SheetBanco.getRange(Linha,151).getValue();
    var	MUN_ACU_2 = SheetBanco.getRange(Linha,152).getValue();
    var	BAIRRO_ACU_2 = SheetBanco.getRange(Linha,153).getValue();
    var	ANTEPOL_ACU_2 = SheetBanco.getRange(Linha,154).getValue();
    var	ANTECRIM_ACU_2 = SheetBanco.getRange(Linha,155).getValue();
    var	ENVTRAF_ACU_2 = SheetBanco.getRange(Linha,156).getValue();
    var	MP_ACU_2 = SheetBanco.getRange(Linha,157).getValue();
    var	MI_ACU_2 = SheetBanco.getRange(Linha,158).getValue();
    var	FOR_ACU_2 = SheetBanco.getRange(Linha,159).getValue();
    
    var HIS_AC_PAI_VIT = SheetBanco.getRange(Linha, 160).getValue();  
    var HIS_AC_MAE_VIT = SheetBanco.getRange(Linha, 161).getValue();
    var VIOL_DOM_MAE_VIT = SheetBanco.getRange(Linha, 162).getValue();
    var MUN_PROF_VIT = SheetBanco.getRange(Linha, 163).getValue();
    var BAIRRO_PROF_VIT = SheetBanco.getRange(Linha, 164).getValue();
    var HIS_MI_VIT = SheetBanco.getRange(Linha, 165).getValue();
    
    var HIS_AC_PAI_ACU_1 = SheetBanco.getRange(Linha, 166).getValue();
    var HIS_AC_MAE_ACU_1 = SheetBanco.getRange(Linha, 167).getValue();
    var VIOL_DOM_MAE_ACU_1 = SheetBanco.getRange(Linha, 168).getValue();
    var HIS_MI_ACU_1 = SheetBanco.getRange(Linha, 169).getValue();
    
    var HIS_AC_PAI_ACU_2 = SheetBanco.getRange(Linha, 170).getValue();
    var HIS_AC_MAE_ACU_2 = SheetBanco.getRange(Linha, 171).getValue();
    var VIOL_DOM_MAE_ACU_2 = SheetBanco.getRange(Linha, 172).getValue();
    var HIS_MI_ACU_2 = SheetBanco.getRange(Linha, 173).getValue();
    
    var CADAVER = SheetBanco.getRange(Linha, 174).getValue();
    var INFORMANTE = SheetBanco.getRange(Linha, 175).getValue();
    var N_INFORMANTE = SheetBanco.getRange(Linha, 176).getValue();
    var OBS = SheetBanco.getRange(Linha, 177).getValue();
    var ALCUNHA_ACU_1 = SheetBanco.getRange(Linha, 178).getValue();
    var ALCUNHA_ACU_2 = SheetBanco.getRange(Linha, 179).getValue();
    var ALCUNHA_VIT = SheetBanco.getRange(Linha, 180).getValue();
    
    Entrada.getRange('B2').activate();
    Entrada.getCurrentCell().setValue(NC);

    Entrada.getRange('B3').activate();
    Entrada.getCurrentCell().setValue(NUM_OCOR);

    Entrada.getRange('B4').activate();
    Entrada.getCurrentCell().setValue(NUM_FATO);

    Entrada.getRange('B5').activate();
    Entrada.getCurrentCell().setValue(N_FATO);

    Entrada.getRange('B6').activate();
    Entrada.getCurrentCell().setValue(NUM_ORG_REGIST);

    Entrada.getRange('B7').activate();
    Entrada.getCurrentCell().setValue(N_ORG_REGIST);

    Entrada.getRange('B8').activate();
    Entrada.getCurrentCell().setValue(ORIGEM_COM);

    Entrada.getRange('B9').activate();
    Entrada.getCurrentCell().setValue(DATA_COM);

    Entrada.getRange('B10').activate();
    Entrada.getCurrentCell().setValue(HORA_COM);

    Entrada.getRange('B11').activate();
    Entrada.getCurrentCell().setValue(ANO_REGIST);

    Entrada.getRange('B12').activate();
    Entrada.getCurrentCell().setValue(FLAGRANTE);

    Entrada.getRange('B13').activate();
    Entrada.getCurrentCell().setValue(TENTATIVA);

    Entrada.getRange('B14').activate();
    Entrada.getCurrentCell().setValue(COMUM);

    Entrada.getRange('B15').activate();
    Entrada.getCurrentCell().setValue(COND);

    Entrada.getRange('B16').activate();
    Entrada.getCurrentCell().setValue(LAUDO);

    Entrada.getRange('B17').activate();
    Entrada.getCurrentCell().setValue(NUM_VIT);

    Entrada.getRange('B18').activate();
    Entrada.getCurrentCell().setValue(NUM_VIT_MORTAS);

    Entrada.getRange('B19').activate();
    Entrada.getCurrentCell().setValue(NUMTEST);

    Entrada.getRange('B20').activate();
    Entrada.getCurrentCell().setValue(TIPO_TEST);

    Entrada.getRange('D3').activate();
    Entrada.getCurrentCell().setValue(DATA_FATO);

    Entrada.getRange('D6').activate();
    Entrada.getCurrentCell().setValue(H_FATO);

    Entrada.getRange('D8').activate();
    Entrada.getCurrentCell().setValue(MUN);

    Entrada.getRange('D19').activate();
    Entrada.getCurrentCell().setValue(A_MUN);

    Entrada.getRange('D20').activate();
    Entrada.getCurrentCell().setValue(LOG);

    Entrada.getRange('F2').activate();
    Entrada.getCurrentCell().setValue(NUM);

    Entrada.getRange('F3').activate();
    Entrada.getCurrentCell().setValue(COMP);

    Entrada.getRange('F4').activate();
    Entrada.getCurrentCell().setValue(CEP);

    Entrada.getRange('F5').activate();
    Entrada.getCurrentCell().setValue(BAIRRO);

    Entrada.getRange('F6').activate();
    Entrada.getCurrentCell().setValue(LAT);

    Entrada.getRange('F7').activate();
    Entrada.getCurrentCell().setValue(LNG);

    Entrada.getRange('F8').activate();
    Entrada.getCurrentCell().setValue(P_REF);

    Entrada.getRange('F9').activate();
    Entrada.getCurrentCell().setValue(TIPO_LOCAL);

    Entrada.getRange('F10').activate();
    Entrada.getCurrentCell().setValue(M_UTIL_AUT);

    Entrada.getRange('F11').activate();
    Entrada.getCurrentCell().setValue(RECUR);

    Entrada.getRange('F12').activate();
    Entrada.getCurrentCell().setValue(NUM_AGRE);

    Entrada.getRange('F13').activate();
    Entrada.getCurrentCell().setValue(M_AUTOR);

    Entrada.getRange('F14').activate();
    Entrada.getCurrentCell().setValue(OBJ);

    Entrada.getRange('F15').activate();
    Entrada.getCurrentCell().setValue(ADTNT);

    Entrada.getRange('F16').activate();
    Entrada.getCurrentCell().setValue(MOT);

    Entrada.getRange('F17').activate();
    Entrada.getCurrentCell().setValue(R_VIT_AUT);

    Entrada.getRange('F18').activate();
    Entrada.getCurrentCell().setValue(MID);

    Entrada.getRange('F19').activate();
    Entrada.getCurrentCell().setValue(HIST_1);

    Entrada.getRange('H19').activate();
    Entrada.getCurrentCell().setValue(HIST_2);

    Entrada.getRange('B22').activate();
    Entrada.getCurrentCell().setValue(N_VIT);

    Entrada.getRange('B23').activate();
    Entrada.getCurrentCell().setValue(RG_VIT);

    Entrada.getRange('B24').activate();
    Entrada.getCurrentCell().setValue(CPF_VIT);

    Entrada.getRange('B25').activate();
    Entrada.getCurrentCell().setValue(SX_VIT);

    Entrada.getRange('B26').activate();
    Entrada.getCurrentCell().setValue(DN_VIT);

    Entrada.getRange('B30').activate();
    Entrada.getCurrentCell().setValue(ORIEN_SX_VIT);

    Entrada.getRange('B31').activate();
    Entrada.getCurrentCell().setValue(COR_VIT);

    Entrada.getRange('B32').activate();
    Entrada.getCurrentCell().setValue(N_PAIVIT);

    Entrada.getRange('B33').activate();
    Entrada.getCurrentCell().setValue(N_MAEVIT);

    Entrada.getRange('B34').activate();
    Entrada.getCurrentCell().setValue(EC_VIT);

    Entrada.getRange('B35').activate();
    Entrada.getCurrentCell().setValue(UE_VIT);

    Entrada.getRange('B36').activate();
    Entrada.getCurrentCell().setValue(N_FIL_VIT);

    Entrada.getRange('B37').activate();
    Entrada.getCurrentCell().setValue(NAT_VIT);

    Entrada.getRange('D22').activate();
    Entrada.getCurrentCell().setValue(NAC_VIT);

    Entrada.getRange('D23').activate();
    Entrada.getCurrentCell().setValue(COND_FIS_VIT);

    Entrada.getRange('D24').activate();
    Entrada.getCurrentCell().setValue(ESC_VIT);

    Entrada.getRange('D25').activate();
    Entrada.getCurrentCell().setValue(ESC_MAEVIT);

    Entrada.getRange('D26').activate();
    Entrada.getCurrentCell().setValue(ESC_PAIVIT);

    Entrada.getRange('D27').activate();
    Entrada.getCurrentCell().setValue(PROF_VIT);

    Entrada.getRange('D28').activate();
    Entrada.getCurrentCell().setValue(END_PROFVIT);

    Entrada.getRange('D29').activate();
    Entrada.getCurrentCell().setValue(END_RESVIT);

    Entrada.getRange('D30').activate();
    Entrada.getCurrentCell().setValue(MUN_VIT);

    Entrada.getRange('D31').activate();
    Entrada.getCurrentCell().setValue(BAIRRO_VIT);

    Entrada.getRange('D32').activate();
    Entrada.getCurrentCell().setValue(ANTEPOL_VIT);

    Entrada.getRange('D33').activate();
    Entrada.getCurrentCell().setValue(ANTECRIM_VIT);

    Entrada.getRange('D34').activate();
    Entrada.getCurrentCell().setValue(ENVTRAF_VIT);

    Entrada.getRange('D35').activate();
    Entrada.getCurrentCell().setValue(MP_VIT);

    Entrada.getRange('D36').activate();
    Entrada.getCurrentCell().setValue(MI_VIT);

    Entrada.getRange('D37').activate();
    Entrada.getCurrentCell().setValue(FOR_VIT);

    Entrada.getRange('B39').activate();
    Entrada.getCurrentCell().setValue(N_ACU_1);

    Entrada.getRange('B40').activate();
    Entrada.getCurrentCell().setValue(RG_ACU_1);

    Entrada.getRange('B41').activate();
    Entrada.getCurrentCell().setValue(CPF_ACU_1);

    Entrada.getRange('B42').activate();
    Entrada.getCurrentCell().setValue(SX_ACU_1);

    Entrada.getRange('B43').activate();
    Entrada.getCurrentCell().setValue(DN_ACU_1);

    Entrada.getRange('B47').activate();
    Entrada.getCurrentCell().setValue(ORIEN_SX_ACU_1);

    Entrada.getRange('B48').activate();
    Entrada.getCurrentCell().setValue(COR_ACU_1);

    Entrada.getRange('B49').activate();
    Entrada.getCurrentCell().setValue(N_PAI_ACU_1);

    Entrada.getRange('B50').activate();
    Entrada.getCurrentCell().setValue(N_MAE_ACU_1);

    Entrada.getRange('B51').activate();
    Entrada.getCurrentCell().setValue(EC_ACU_1);

    Entrada.getRange('B52').activate();
    Entrada.getCurrentCell().setValue(UE_ACU_1);

    Entrada.getRange('B53').activate();
    Entrada.getCurrentCell().setValue(N_FIL_ACU_1);

    Entrada.getRange('B54').activate();
    Entrada.getCurrentCell().setValue(NAT_ACU_1);

    Entrada.getRange('D39').activate();
    Entrada.getCurrentCell().setValue(NAC_ACU_1);

    Entrada.getRange('D40').activate();
    Entrada.getCurrentCell().setValue(COND_FIS_ACU_1);

    Entrada.getRange('D41').activate();
    Entrada.getCurrentCell().setValue(ESC_ACU_1);

    Entrada.getRange('D42').activate();
    Entrada.getCurrentCell().setValue(ESC_MAE_ACU_1);
    
    Entrada.getRange('D43').activate();
    Entrada.getCurrentCell().setValue(ESC_PAI_ACU_1);

    Entrada.getRange('D44').activate();
    Entrada.getCurrentCell().setValue(PROF_ACU_1);

    Entrada.getRange('D45').activate();
    Entrada.getCurrentCell().setValue(END_PROF_ACU_1);

    Entrada.getRange('D46').activate();
    Entrada.getCurrentCell().setValue(END_RES_ACU_1);

    Entrada.getRange('D47').activate();
    Entrada.getCurrentCell().setValue(MUN_ACU_1);

    Entrada.getRange('D48').activate();
    Entrada.getCurrentCell().setValue(BAIRRO_ACU_1);

    Entrada.getRange('D49').activate();
    Entrada.getCurrentCell().setValue(ANTEPOL_ACU_1);

    Entrada.getRange('D50').activate();
    Entrada.getCurrentCell().setValue(ANTECRIM_ACU_1);

    Entrada.getRange('D51').activate();
    Entrada.getCurrentCell().setValue(ENVTRAF_ACU_1);

    Entrada.getRange('D52').activate();
    Entrada.getCurrentCell().setValue(MP_ACU_1);

    Entrada.getRange('D53').activate();
    Entrada.getCurrentCell().setValue(MI_ACU_1);

    Entrada.getRange('D54').activate();
    Entrada.getCurrentCell().setValue(FOR_ACU_1);

    Entrada.getRange('F39').activate();
    Entrada.getCurrentCell().setValue(N_ACU_2);

    Entrada.getRange('F40').activate();
    Entrada.getCurrentCell().setValue(RG_ACU_2);

    Entrada.getRange('F41').activate();
    Entrada.getCurrentCell().setValue(CPF_ACU_2);

    Entrada.getRange('F42').activate();
    Entrada.getCurrentCell().setValue(SX_ACU_2);

    Entrada.getRange('F43').activate();
    Entrada.getCurrentCell().setValue(DN_ACU_2);

    Entrada.getRange('F47').activate();
    Entrada.getCurrentCell().setValue(ORIEN_SX_ACU_2);

    Entrada.getRange('F48').activate();
    Entrada.getCurrentCell().setValue(COR_ACU_2);

    Entrada.getRange('F49').activate();
    Entrada.getCurrentCell().setValue(N_PAI_ACU_2);

    Entrada.getRange('F50').activate();
    Entrada.getCurrentCell().setValue(N_MAE_ACU_2);

    Entrada.getRange('F51').activate();
    Entrada.getCurrentCell().setValue(EC_ACU_2);

    Entrada.getRange('F52').activate();
    Entrada.getCurrentCell().setValue(UE_ACU_2);

    Entrada.getRange('F53').activate();
    Entrada.getCurrentCell().setValue(N_FIL_ACU_2);

    Entrada.getRange('F54').activate();
    Entrada.getCurrentCell().setValue(NAT_ACU_2);

    Entrada.getRange('H39').activate();
    Entrada.getCurrentCell().setValue(NAC_ACU_2);

    Entrada.getRange('H40').activate();
    Entrada.getCurrentCell().setValue(COND_FIS_ACU_2);

    Entrada.getRange('H41').activate();
    Entrada.getCurrentCell().setValue(ESC_ACU_2);

    Entrada.getRange('H42').activate();
    Entrada.getCurrentCell().setValue(ESC_MAE_ACU_2);

    Entrada.getRange('H43').activate();
    Entrada.getCurrentCell().setValue(ESC_PAI_ACU_2);

    Entrada.getRange('H44').activate();
    Entrada.getCurrentCell().setValue(PROF_ACU_2);

    Entrada.getRange('H45').activate();
    Entrada.getCurrentCell().setValue(END_PROF_ACU_2);

    Entrada.getRange('H46').activate();
    Entrada.getCurrentCell().setValue(END_RES_ACU_2);

    Entrada.getRange('H47').activate();
    Entrada.getCurrentCell().setValue(MUN_ACU_2);

    Entrada.getRange('H48').activate();
    Entrada.getCurrentCell().setValue(BAIRRO_ACU_2);

    Entrada.getRange('H49').activate();
    Entrada.getCurrentCell().setValue(ANTEPOL_ACU_2);

    Entrada.getRange('H50').activate();
    Entrada.getCurrentCell().setValue(ANTECRIM_ACU_2);

    Entrada.getRange('H51').activate();
    Entrada.getCurrentCell().setValue(ENVTRAF_ACU_2);

    Entrada.getRange('H52').activate();
    Entrada.getCurrentCell().setValue(MP_ACU_2);

    Entrada.getRange('H53').activate();
    Entrada.getCurrentCell().setValue(MI_ACU_2);

    Entrada.getRange('H54').activate();
    Entrada.getCurrentCell().setValue(FOR_ACU_2);
  
    Entrada.getRange('H21').activate();
    Entrada.getCurrentCell().setValue(OBS);
    
    Entrada.getRange('F24').activate();
    Entrada.getCurrentCell().setValue(HIS_AC_PAI_VIT);	

    Entrada.getRange('F25').activate();
    Entrada.getCurrentCell().setValue(HIS_AC_MAE_VIT);

    Entrada.getRange('F26').activate();
    Entrada.getCurrentCell().setValue(VIOL_DOM_MAE_VIT);

    Entrada.getRange('F27').activate();
    Entrada.getCurrentCell().setValue(MUN_PROF_VIT);

    Entrada.getRange('F28').activate();
    Entrada.getCurrentCell().setValue(BAIRRO_PROF_VIT);

    Entrada.getRange('F36').activate();
    Entrada.getCurrentCell().setValue(HIS_MI_VIT);

    Entrada.getRange('D56').activate();
    Entrada.getCurrentCell().setValue(HIS_AC_PAI_ACU_1);

    Entrada.getRange('D57').activate();
    Entrada.getCurrentCell().setValue(HIS_AC_MAE_ACU_1);

    Entrada.getRange('D58').activate();
    Entrada.getCurrentCell().setValue(VIOL_DOM_MAE_ACU_1);

    Entrada.getRange('D55').activate();
    Entrada.getCurrentCell().setValue(HIS_MI_ACU_1);

    Entrada.getRange('H56').activate();
    Entrada.getCurrentCell().setValue(HIS_AC_PAI_ACU_2);

    Entrada.getRange('H57').activate();
    Entrada.getCurrentCell().setValue(HIS_AC_MAE_ACU_2);

    Entrada.getRange('H58').activate();
    Entrada.getCurrentCell().setValue(VIOL_DOM_MAE_ACU_2);

    Entrada.getRange('H55').activate();
    Entrada.getCurrentCell().setValue(HIS_MI_ACU_2);

    Entrada.getRange('H2').activate();
    Entrada.getCurrentCell().setValue(CADAVER);

    Entrada.getRange('H3').activate();
    Entrada.getCurrentCell().setValue(INFORMANTE);
 
    Entrada.getRange('H4').activate();
    Entrada.getCurrentCell().setValue(N_INFORMANTE);
    
    Entrada.getRange('B55').activate();
    Entrada.getCurrentCell().setValue(ALCUNHA_ACU_1);
    
    Entrada.getRange('F55').activate();
    Entrada.getCurrentCell().setValue(ALCUNHA_ACU_2);
    
    Entrada.getRange('F22').activate();
    Entrada.getCurrentCell().setValue(ALCUNHA_VIT);

  }else{
   Browser.msgBox("Não localizado!") 
  
  }

}


function EditarOcorrencia(){ 
  
  var Entrada = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Entrada');
  var Pesquisa = Entrada.getRange('B2').getValue();
  var SheetBanco = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Banco');
  SheetBanco.getActiveCell();
  
  var LocalPesquisa = SheetBanco.getRange(2, 2, SheetBanco.getLastRow()).getValues(); 
  var Resultado = LocalPesquisa.Pesquisa(Pesquisa);
  
  var Linha = Resultado + 2;
  var verif = Verificar();
  
  if(Resultado != -1){
   if(verif == true){
     SheetBanco.getActiveCell();
    
    SheetBanco.getRange(Linha,2).setValue(Entrada.getRange('B2').getValue());
    SheetBanco.getRange(Linha,4).setValue(Entrada.getRange('B3').getValue());
    SheetBanco.getRange(Linha,3).setValue(Entrada.getRange('B4').getValue());
    SheetBanco.getRange(Linha,5).setValue(Entrada.getRange('B5').getValue());
    SheetBanco.getRange(Linha,6).setValue(Entrada.getRange('B6').getValue());
    SheetBanco.getRange(Linha,7).setValue(Entrada.getRange('B7').getValue());
    SheetBanco.getRange(Linha,8).setValue(Entrada.getRange('B8').getValue());
    SheetBanco.getRange(Linha,9).setValue(Entrada.getRange('B9').getValue());
    SheetBanco.getRange(Linha,10).setValue(Entrada.getRange('B10').getValue());
    SheetBanco.getRange(Linha,11).setValue(Entrada.getRange('B11').getValue());
    SheetBanco.getRange(Linha,12).setValue(Entrada.getRange('B12').getValue());
    SheetBanco.getRange(Linha,13).setValue(Entrada.getRange('B13').getValue());
    SheetBanco.getRange(Linha,14).setValue(Entrada.getRange('B14').getValue());
    SheetBanco.getRange(Linha,15).setValue(Entrada.getRange('B15').getValue());
    SheetBanco.getRange(Linha,16).setValue(Entrada.getRange('B16').getValue());
    SheetBanco.getRange(Linha,17).setValue(Entrada.getRange('B17').getValue());
    SheetBanco.getRange(Linha,18).setValue(Entrada.getRange('B18').getValue());
    SheetBanco.getRange(Linha,19).setValue(Entrada.getRange('B19').getValue());
    SheetBanco.getRange(Linha,20).setValue(Entrada.getRange('B20').getValue());
    SheetBanco.getRange(Linha,21).setValue(Entrada.getRange('D2').getValue());
    SheetBanco.getRange(Linha,22).setValue(Entrada.getRange('D3').getValue());
    SheetBanco.getRange(Linha,24).setValue(Entrada.getRange('D4').getValue());
    SheetBanco.getRange(Linha,25).setValue(Entrada.getRange('D5').getValue());
    SheetBanco.getRange(Linha,27).setValue(Entrada.getRange('D6').getValue());
    SheetBanco.getRange(Linha,28).setValue(Entrada.getRange('D7').getValue());
    SheetBanco.getRange(Linha,29).setValue(Entrada.getRange('D8').getValue());
    SheetBanco.getRange(Linha,30).setValue(Entrada.getRange('D9').getValue());
    SheetBanco.getRange(Linha,31).setValue(Entrada.getRange('D10').getValue());
    SheetBanco.getRange(Linha,32).setValue(Entrada.getRange('D11').getValue());
    SheetBanco.getRange(Linha,33).setValue(Entrada.getRange('D12').getValue());
    SheetBanco.getRange(Linha,34).setValue(Entrada.getRange('D13').getValue());
    SheetBanco.getRange(Linha,35).setValue(Entrada.getRange('D14').getValue());
    SheetBanco.getRange(Linha,36).setValue(Entrada.getRange('D15').getValue());
    SheetBanco.getRange(Linha,37).setValue(Entrada.getRange('D16').getValue());
    SheetBanco.getRange(Linha,38).setValue(Entrada.getRange('D17').getValue());
    SheetBanco.getRange(Linha,39).setValue(Entrada.getRange('D18').getValue());
    SheetBanco.getRange(Linha,40).setValue(Entrada.getRange('D19').getValue());
    SheetBanco.getRange(Linha,41).setValue(Entrada.getRange('D20').getValue());
    SheetBanco.getRange(Linha,42).setValue(Entrada.getRange('F2').getValue());
    SheetBanco.getRange(Linha,43).setValue(Entrada.getRange('F3').getValue());
    SheetBanco.getRange(Linha,44).setValue(Entrada.getRange('F4').getValue());
    SheetBanco.getRange(Linha,45).setValue(Entrada.getRange('F5').getValue());
    SheetBanco.getRange(Linha,46).setValue(Entrada.getRange('F6').getValue());
    SheetBanco.getRange(Linha,47).setValue(Entrada.getRange('F7').getValue());
    SheetBanco.getRange(Linha,48).setValue(Entrada.getRange('F8').getValue());
    SheetBanco.getRange(Linha,49).setValue(Entrada.getRange('F9').getValue());
    SheetBanco.getRange(Linha,50).setValue(Entrada.getRange('F10').getValue());
    SheetBanco.getRange(Linha,51).setValue(Entrada.getRange('F11').getValue());
    SheetBanco.getRange(Linha,52).setValue(Entrada.getRange('F12').getValue());
    SheetBanco.getRange(Linha,53).setValue(Entrada.getRange('F13').getValue());
    SheetBanco.getRange(Linha,54).setValue(Entrada.getRange('F14').getValue());
    SheetBanco.getRange(Linha,55).setValue(Entrada.getRange('F15').getValue());
    SheetBanco.getRange(Linha,56).setValue(Entrada.getRange('F16').getValue());
    SheetBanco.getRange(Linha,57).setValue(Entrada.getRange('F17').getValue());
    SheetBanco.getRange(Linha,58).setValue(Entrada.getRange('F18').getValue());
    SheetBanco.getRange(Linha,59).setValue(Entrada.getRange('F19').getValue());
    SheetBanco.getRange(Linha,60).setValue(Entrada.getRange('H19').getValue());
    SheetBanco.getRange(Linha,63).setValue(Entrada.getRange('B22').getValue());
    SheetBanco.getRange(Linha,64).setValue(Entrada.getRange('B23').getValue());
    SheetBanco.getRange(Linha,65).setValue(Entrada.getRange('B24').getValue());
    SheetBanco.getRange(Linha,66).setValue(Entrada.getRange('B25').getValue());
    SheetBanco.getRange(Linha,67).setValue(Entrada.getRange('B26').getValue());
    SheetBanco.getRange(Linha,68).setValue(Entrada.getRange('B27').getValue());
    SheetBanco.getRange(Linha,69).setValue(Entrada.getRange('B28').getValue());
    SheetBanco.getRange(Linha,70).setValue(Entrada.getRange('B29').getValue());
    SheetBanco.getRange(Linha,71).setValue(Entrada.getRange('B30').getValue());
    SheetBanco.getRange(Linha,72).setValue(Entrada.getRange('B31').getValue());
    SheetBanco.getRange(Linha,73).setValue(Entrada.getRange('B32').getValue());
    SheetBanco.getRange(Linha,74).setValue(Entrada.getRange('B33').getValue());
    SheetBanco.getRange(Linha,75).setValue(Entrada.getRange('B34').getValue());
    SheetBanco.getRange(Linha,76).setValue(Entrada.getRange('B35').getValue());
    SheetBanco.getRange(Linha,77).setValue(Entrada.getRange('B36').getValue());
    SheetBanco.getRange(Linha,78).setValue(Entrada.getRange('B37').getValue());
    SheetBanco.getRange(Linha,79).setValue(Entrada.getRange('D22').getValue());
    SheetBanco.getRange(Linha,80).setValue(Entrada.getRange('D23').getValue());
    SheetBanco.getRange(Linha,81).setValue(Entrada.getRange('D24').getValue());
    SheetBanco.getRange(Linha,82).setValue(Entrada.getRange('D25').getValue());
    SheetBanco.getRange(Linha,83).setValue(Entrada.getRange('D26').getValue());
    SheetBanco.getRange(Linha,85).setValue(Entrada.getRange('D27').getValue());
    SheetBanco.getRange(Linha,86).setValue(Entrada.getRange('D28').getValue());
    SheetBanco.getRange(Linha,87).setValue(Entrada.getRange('D29').getValue());
    SheetBanco.getRange(Linha,88).setValue(Entrada.getRange('D30').getValue());
    SheetBanco.getRange(Linha,89).setValue(Entrada.getRange('D31').getValue());
    SheetBanco.getRange(Linha,90).setValue(Entrada.getRange('D32').getValue());
    SheetBanco.getRange(Linha,91).setValue(Entrada.getRange('D33').getValue());
    SheetBanco.getRange(Linha,92).setValue(Entrada.getRange('D34').getValue());
    SheetBanco.getRange(Linha,93).setValue(Entrada.getRange('D35').getValue());
    SheetBanco.getRange(Linha,94).setValue(Entrada.getRange('D36').getValue());
    SheetBanco.getRange(Linha,95).setValue(Entrada.getRange('D37').getValue());
    SheetBanco.getRange(Linha,96).setValue(Entrada.getRange('B39').getValue());
    SheetBanco.getRange(Linha,97).setValue(Entrada.getRange('B40').getValue());
    SheetBanco.getRange(Linha,98).setValue(Entrada.getRange('B41').getValue());
    SheetBanco.getRange(Linha,99).setValue(Entrada.getRange('B42').getValue());
    SheetBanco.getRange(Linha,100).setValue(Entrada.getRange('B43').getValue());
    SheetBanco.getRange(Linha,101).setValue(Entrada.getRange('B44').getValue());
    SheetBanco.getRange(Linha,102).setValue(Entrada.getRange('B45').getValue());
    SheetBanco.getRange(Linha,103).setValue(Entrada.getRange('B46').getValue());
    SheetBanco.getRange(Linha,104).setValue(Entrada.getRange('B47').getValue());
    SheetBanco.getRange(Linha,105).setValue(Entrada.getRange('B48').getValue());
    SheetBanco.getRange(Linha,106).setValue(Entrada.getRange('B49').getValue());
    SheetBanco.getRange(Linha,107).setValue(Entrada.getRange('B50').getValue());
    SheetBanco.getRange(Linha,108).setValue(Entrada.getRange('B51').getValue());
    SheetBanco.getRange(Linha,109).setValue(Entrada.getRange('B52').getValue());
    SheetBanco.getRange(Linha,110).setValue(Entrada.getRange('B53').getValue());
    SheetBanco.getRange(Linha,111).setValue(Entrada.getRange('B54').getValue());
    SheetBanco.getRange(Linha,112).setValue(Entrada.getRange('D39').getValue());
    SheetBanco.getRange(Linha,113).setValue(Entrada.getRange('D40').getValue());
    SheetBanco.getRange(Linha,114).setValue(Entrada.getRange('D41').getValue());
    SheetBanco.getRange(Linha,115).setValue(Entrada.getRange('D42').getValue());
    SheetBanco.getRange(Linha,116).setValue(Entrada.getRange('D43').getValue());
    SheetBanco.getRange(Linha,117).setValue(Entrada.getRange('D44').getValue());
    SheetBanco.getRange(Linha,118).setValue(Entrada.getRange('D45').getValue());
    SheetBanco.getRange(Linha,119).setValue(Entrada.getRange('D46').getValue());
    SheetBanco.getRange(Linha,120).setValue(Entrada.getRange('D47').getValue());
    SheetBanco.getRange(Linha,121).setValue(Entrada.getRange('D48').getValue());
    SheetBanco.getRange(Linha,122).setValue(Entrada.getRange('D49').getValue());
    SheetBanco.getRange(Linha,123).setValue(Entrada.getRange('D50').getValue());
    SheetBanco.getRange(Linha,124).setValue(Entrada.getRange('D51').getValue());
    SheetBanco.getRange(Linha,125).setValue(Entrada.getRange('D52').getValue());
    SheetBanco.getRange(Linha,126).setValue(Entrada.getRange('D53').getValue());
    SheetBanco.getRange(Linha,127).setValue(Entrada.getRange('D54').getValue());
    SheetBanco.getRange(Linha,128).setValue(Entrada.getRange('F39').getValue());
    SheetBanco.getRange(Linha,129).setValue(Entrada.getRange('F40').getValue());
    SheetBanco.getRange(Linha,130).setValue(Entrada.getRange('F41').getValue());
    SheetBanco.getRange(Linha,131).setValue(Entrada.getRange('F42').getValue());
    SheetBanco.getRange(Linha,132).setValue(Entrada.getRange('F43').getValue());
    SheetBanco.getRange(Linha,133).setValue(Entrada.getRange('F44').getValue());
    SheetBanco.getRange(Linha,134).setValue(Entrada.getRange('F45').getValue());
    SheetBanco.getRange(Linha,135).setValue(Entrada.getRange('F46').getValue());
    SheetBanco.getRange(Linha,136).setValue(Entrada.getRange('F47').getValue());
    SheetBanco.getRange(Linha,137).setValue(Entrada.getRange('F48').getValue());
    SheetBanco.getRange(Linha,138).setValue(Entrada.getRange('F49').getValue());
    SheetBanco.getRange(Linha,139).setValue(Entrada.getRange('F50').getValue());
    SheetBanco.getRange(Linha,140).setValue(Entrada.getRange('F51').getValue());
    SheetBanco.getRange(Linha,141).setValue(Entrada.getRange('F52').getValue());
    SheetBanco.getRange(Linha,142).setValue(Entrada.getRange('F53').getValue());
    SheetBanco.getRange(Linha,143).setValue(Entrada.getRange('F54').getValue());
    SheetBanco.getRange(Linha,144).setValue(Entrada.getRange('H39').getValue());
    SheetBanco.getRange(Linha,145).setValue(Entrada.getRange('H40').getValue());
    SheetBanco.getRange(Linha,146).setValue(Entrada.getRange('H41').getValue());
    SheetBanco.getRange(Linha,147).setValue(Entrada.getRange('H42').getValue());
    SheetBanco.getRange(Linha,148).setValue(Entrada.getRange('H43').getValue());
    SheetBanco.getRange(Linha,149).setValue(Entrada.getRange('H44').getValue());
    SheetBanco.getRange(Linha,150).setValue(Entrada.getRange('H45').getValue());
    SheetBanco.getRange(Linha,151).setValue(Entrada.getRange('H46').getValue());
    SheetBanco.getRange(Linha,152).setValue(Entrada.getRange('H47').getValue());
    SheetBanco.getRange(Linha,153).setValue(Entrada.getRange('H48').getValue());
    SheetBanco.getRange(Linha,154).setValue(Entrada.getRange('H49').getValue());
    SheetBanco.getRange(Linha,155).setValue(Entrada.getRange('H50').getValue());
    SheetBanco.getRange(Linha,156).setValue(Entrada.getRange('H51').getValue());
    SheetBanco.getRange(Linha,157).setValue(Entrada.getRange('H52').getValue());
    SheetBanco.getRange(Linha,158).setValue(Entrada.getRange('H53').getValue());
    SheetBanco.getRange(Linha,159).setValue(Entrada.getRange('H54').getValue());
     
    SheetBanco.getRange(Linha,160).setValue(Entrada.getRange('F24').getValue());
    SheetBanco.getRange(Linha,161).setValue(Entrada.getRange('F25').getValue());
    SheetBanco.getRange(Linha,162).setValue(Entrada.getRange('F26').getValue());
    SheetBanco.getRange(Linha,163).setValue(Entrada.getRange('F27').getValue());
    SheetBanco.getRange(Linha,164).setValue(Entrada.getRange('F28').getValue());
    SheetBanco.getRange(Linha,165).setValue(Entrada.getRange('F36').getValue());
    
    SheetBanco.getRange(Linha,166).setValue(Entrada.getRange('D56').getValue());
    SheetBanco.getRange(Linha,167).setValue(Entrada.getRange('D57').getValue());
    SheetBanco.getRange(Linha,168).setValue(Entrada.getRange('D58').getValue());
    SheetBanco.getRange(Linha,169).setValue(Entrada.getRange('D55').getValue());

    SheetBanco.getRange(Linha,170).setValue(Entrada.getRange('H56').getValue());
    SheetBanco.getRange(Linha,171).setValue(Entrada.getRange('H57').getValue());
    SheetBanco.getRange(Linha,172).setValue(Entrada.getRange('H58').getValue());
    SheetBanco.getRange(Linha,173).setValue(Entrada.getRange('H55').getValue());
     
    SheetBanco.getRange(Linha,174).setValue(Entrada.getRange('H2').getValue());
    SheetBanco.getRange(Linha,175).setValue(Entrada.getRange('H3').getValue());
    SheetBanco.getRange(Linha,176).setValue(Entrada.getRange('H4').getValue());
     
    SheetBanco.getRange(Linha,177).setValue(Entrada.getRange('H21').getValue());
     
    SheetBanco.getRange(Linha,178).setValue(Entrada.getRange('B55').getValue());
     
    SheetBanco.getRange(Linha,179).setValue(Entrada.getRange('F55').getValue());
     
    SheetBanco.getRange(Linha,179).setValue(Entrada.getRange('F22').getValue());
    
     Browser.msgBox('Ocorrência Editada!')
    
    Entrada.getActiveCell();
    
    LimparEntrada();
    
    Entrada.getRange('B2').activate();
                    }else{
    Browser.msgBox("Preencha todos os campos!")
                    }
  } else {
    
    Browser.msgBox('Ocorrência não localizada!')
    
  }

}
                   


function LocalizarCelulaVazia(){
  
  var SheetBanco = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Banco');
  //var Entrada = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Entrada');
  
  SheetBanco.getActiveCell();
  SheetBanco.getRange('A1').activate();
  
  SheetBanco.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate();
  
  var Linha = SheetBanco.getCurrentCell().getRow();
  var NL = SheetBanco.getRange(Linha, 1).getValue();
  
  NL = NL + 1
  
  SheetBanco.getActiveCell().offset(1,0).activate()
  
  var LinhaR = Linha + 1
  
  //Entrada.getRange(2,2).setValue(NL);
  SheetBanco.getRange(LinhaR, 1).setValue(NL);
  
}

function LocalizarCelulaVazia_Excluida(){
  
  var Excluidas = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Excluidas');
   
  Excluidas.getActiveCell();
  Excluidas.getRange('A1').activate();
  
  Excluidas.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate();
  
  var Linha = Excluidas.getCurrentCell().getRow();
  var NL2 = Excluidas.getRange(Linha, 1).getValue();
  
  NL2 = NL2 + 1
  
  Excluidas.getActiveCell().offset(1,0).activate()
  
  var LinhaE = Linha + 1
  
  //Entrada.getRange(2,2).setValue(NL);
  Excluidas.getRange(LinhaE, 1).setValue(NL2);
  
 
}

function SalvarDados(){
  
  var Entrada = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Entrada'); 
  
  if(Entrada.getRange(2, 2).getValue() == ""){
    Browser.msgBox("Preencher campo Nº CONTROLE");
    Entrada.getRange('B2').activate();
    return false
  }
  
  var verif = Verificar();
  
  if(verif == true){
    LocalizarCelulaVazia();
  }else{
    Browser.msgBox("Preencha todos os campos!")
  }
  
  var SheetBanco = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Banco');
  var Linha = SheetBanco.getCurrentCell().getRow();
  
  var LocalPesquisa = SheetBanco.getRange(2, 2, SheetBanco.getLastRow()).getValues();
  var Resultado = LocalPesquisa.Pesquisa(Entrada.getRange('B2').getValue());
  
  if(Resultado !== -1){
  
    Browser.msgBox("Essa ocorrencia já está cadastrada!")
    Entrada.getRange('B2').activate();
    
    return false
  }
  
    SheetBanco.getRange(Linha,2).setValue(Entrada.getRange('B2').getValue());
    SheetBanco.getRange(Linha,4).setValue(Entrada.getRange('B3').getValue());
    SheetBanco.getRange(Linha,3).setValue(Entrada.getRange('B4').getValue());
    SheetBanco.getRange(Linha,5).setValue(Entrada.getRange('B5').getValue());
    SheetBanco.getRange(Linha,6).setValue(Entrada.getRange('B6').getValue());
    SheetBanco.getRange(Linha,7).setValue(Entrada.getRange('B7').getValue());
    SheetBanco.getRange(Linha,8).setValue(Entrada.getRange('B8').getValue());
    SheetBanco.getRange(Linha,9).setValue(Entrada.getRange('B9').getValue());
    SheetBanco.getRange(Linha,10).setValue(Entrada.getRange('B10').getValue());
    SheetBanco.getRange(Linha,11).setValue(Entrada.getRange('B11').getValue());
    SheetBanco.getRange(Linha,12).setValue(Entrada.getRange('B12').getValue());
    SheetBanco.getRange(Linha,13).setValue(Entrada.getRange('B13').getValue());
    SheetBanco.getRange(Linha,14).setValue(Entrada.getRange('B14').getValue());
    SheetBanco.getRange(Linha,15).setValue(Entrada.getRange('B15').getValue());
    SheetBanco.getRange(Linha,16).setValue(Entrada.getRange('B16').getValue());
    SheetBanco.getRange(Linha,17).setValue(Entrada.getRange('B17').getValue());
    SheetBanco.getRange(Linha,18).setValue(Entrada.getRange('B18').getValue());
    SheetBanco.getRange(Linha,19).setValue(Entrada.getRange('B19').getValue());
    SheetBanco.getRange(Linha,20).setValue(Entrada.getRange('B20').getValue());
    SheetBanco.getRange(Linha,21).setValue(Entrada.getRange('D2').getValue());
    SheetBanco.getRange(Linha,22).setValue(Entrada.getRange('D3').getValue());
    SheetBanco.getRange(Linha,24).setValue(Entrada.getRange('D4').getValue());
    SheetBanco.getRange(Linha,25).setValue(Entrada.getRange('D5').getValue());
    SheetBanco.getRange(Linha,27).setValue(Entrada.getRange('D6').getValue());
    SheetBanco.getRange(Linha,28).setValue(Entrada.getRange('D7').getValue());
    SheetBanco.getRange(Linha,29).setValue(Entrada.getRange('D8').getValue());
    SheetBanco.getRange(Linha,30).setValue(Entrada.getRange('D9').getValue());
    SheetBanco.getRange(Linha,31).setValue(Entrada.getRange('D10').getValue());
    SheetBanco.getRange(Linha,32).setValue(Entrada.getRange('D11').getValue());
    SheetBanco.getRange(Linha,33).setValue(Entrada.getRange('D12').getValue());
    SheetBanco.getRange(Linha,34).setValue(Entrada.getRange('D13').getValue());
    SheetBanco.getRange(Linha,35).setValue(Entrada.getRange('D14').getValue());
    SheetBanco.getRange(Linha,36).setValue(Entrada.getRange('D15').getValue());
    SheetBanco.getRange(Linha,37).setValue(Entrada.getRange('D16').getValue());
    SheetBanco.getRange(Linha,38).setValue(Entrada.getRange('D17').getValue());
    SheetBanco.getRange(Linha,39).setValue(Entrada.getRange('D18').getValue());
    SheetBanco.getRange(Linha,40).setValue(Entrada.getRange('D19').getValue());
    SheetBanco.getRange(Linha,41).setValue(Entrada.getRange('D20').getValue());
    SheetBanco.getRange(Linha,42).setValue(Entrada.getRange('F2').getValue());
    SheetBanco.getRange(Linha,43).setValue(Entrada.getRange('F3').getValue());
    SheetBanco.getRange(Linha,44).setValue(Entrada.getRange('F4').getValue());
    SheetBanco.getRange(Linha,45).setValue(Entrada.getRange('F5').getValue());
    SheetBanco.getRange(Linha,46).setValue(Entrada.getRange('F6').getValue());
    SheetBanco.getRange(Linha,47).setValue(Entrada.getRange('F7').getValue());
    SheetBanco.getRange(Linha,48).setValue(Entrada.getRange('F8').getValue());
    SheetBanco.getRange(Linha,49).setValue(Entrada.getRange('F9').getValue());
    SheetBanco.getRange(Linha,50).setValue(Entrada.getRange('F10').getValue());
    SheetBanco.getRange(Linha,51).setValue(Entrada.getRange('F11').getValue());
    SheetBanco.getRange(Linha,52).setValue(Entrada.getRange('F12').getValue());
    SheetBanco.getRange(Linha,53).setValue(Entrada.getRange('F13').getValue());
    SheetBanco.getRange(Linha,54).setValue(Entrada.getRange('F14').getValue());
    SheetBanco.getRange(Linha,55).setValue(Entrada.getRange('F15').getValue());
    SheetBanco.getRange(Linha,56).setValue(Entrada.getRange('F16').getValue());
    SheetBanco.getRange(Linha,57).setValue(Entrada.getRange('F17').getValue());
    SheetBanco.getRange(Linha,58).setValue(Entrada.getRange('F18').getValue());
    SheetBanco.getRange(Linha,59).setValue(Entrada.getRange('F19').getValue());
    SheetBanco.getRange(Linha,60).setValue(Entrada.getRange('H19').getValue());
    SheetBanco.getRange(Linha,63).setValue(Entrada.getRange('B22').getValue());
    SheetBanco.getRange(Linha,64).setValue(Entrada.getRange('B23').getValue());
    SheetBanco.getRange(Linha,65).setValue(Entrada.getRange('B24').getValue());
    SheetBanco.getRange(Linha,66).setValue(Entrada.getRange('B25').getValue());
    SheetBanco.getRange(Linha,67).setValue(Entrada.getRange('B26').getValue());
    SheetBanco.getRange(Linha,68).setValue(Entrada.getRange('B27').getValue());
    SheetBanco.getRange(Linha,69).setValue(Entrada.getRange('B28').getValue());
    SheetBanco.getRange(Linha,70).setValue(Entrada.getRange('B29').getValue());
    SheetBanco.getRange(Linha,71).setValue(Entrada.getRange('B30').getValue());
    SheetBanco.getRange(Linha,72).setValue(Entrada.getRange('B31').getValue());
    SheetBanco.getRange(Linha,73).setValue(Entrada.getRange('B32').getValue());
    SheetBanco.getRange(Linha,74).setValue(Entrada.getRange('B33').getValue());
    SheetBanco.getRange(Linha,75).setValue(Entrada.getRange('B34').getValue());
    SheetBanco.getRange(Linha,76).setValue(Entrada.getRange('B35').getValue());
    SheetBanco.getRange(Linha,77).setValue(Entrada.getRange('B36').getValue());
    SheetBanco.getRange(Linha,78).setValue(Entrada.getRange('B37').getValue());
    SheetBanco.getRange(Linha,79).setValue(Entrada.getRange('D22').getValue());
    SheetBanco.getRange(Linha,80).setValue(Entrada.getRange('D23').getValue());
    SheetBanco.getRange(Linha,81).setValue(Entrada.getRange('D24').getValue());
    SheetBanco.getRange(Linha,82).setValue(Entrada.getRange('D25').getValue());
    SheetBanco.getRange(Linha,83).setValue(Entrada.getRange('D26').getValue());
    SheetBanco.getRange(Linha,85).setValue(Entrada.getRange('D27').getValue());
    SheetBanco.getRange(Linha,86).setValue(Entrada.getRange('D28').getValue());
    SheetBanco.getRange(Linha,87).setValue(Entrada.getRange('D29').getValue());
    SheetBanco.getRange(Linha,88).setValue(Entrada.getRange('D30').getValue());
    SheetBanco.getRange(Linha,89).setValue(Entrada.getRange('D31').getValue());
    SheetBanco.getRange(Linha,90).setValue(Entrada.getRange('D32').getValue());
    SheetBanco.getRange(Linha,91).setValue(Entrada.getRange('D33').getValue());
    SheetBanco.getRange(Linha,92).setValue(Entrada.getRange('D34').getValue());
    SheetBanco.getRange(Linha,93).setValue(Entrada.getRange('D35').getValue());
    SheetBanco.getRange(Linha,94).setValue(Entrada.getRange('D36').getValue());
    SheetBanco.getRange(Linha,95).setValue(Entrada.getRange('D37').getValue());
    SheetBanco.getRange(Linha,96).setValue(Entrada.getRange('B39').getValue());
    SheetBanco.getRange(Linha,97).setValue(Entrada.getRange('B40').getValue());
    SheetBanco.getRange(Linha,98).setValue(Entrada.getRange('B41').getValue());
    SheetBanco.getRange(Linha,99).setValue(Entrada.getRange('B42').getValue());
    SheetBanco.getRange(Linha,100).setValue(Entrada.getRange('B43').getValue());
    SheetBanco.getRange(Linha,101).setValue(Entrada.getRange('B44').getValue());
    SheetBanco.getRange(Linha,102).setValue(Entrada.getRange('B45').getValue());
    SheetBanco.getRange(Linha,103).setValue(Entrada.getRange('B46').getValue());
    SheetBanco.getRange(Linha,104).setValue(Entrada.getRange('B47').getValue());
    SheetBanco.getRange(Linha,105).setValue(Entrada.getRange('B48').getValue());
    SheetBanco.getRange(Linha,106).setValue(Entrada.getRange('B49').getValue());
    SheetBanco.getRange(Linha,107).setValue(Entrada.getRange('B50').getValue());
    SheetBanco.getRange(Linha,108).setValue(Entrada.getRange('B51').getValue());
    SheetBanco.getRange(Linha,109).setValue(Entrada.getRange('B52').getValue());
    SheetBanco.getRange(Linha,110).setValue(Entrada.getRange('B53').getValue());
    SheetBanco.getRange(Linha,111).setValue(Entrada.getRange('B54').getValue());
    SheetBanco.getRange(Linha,112).setValue(Entrada.getRange('D39').getValue());
    SheetBanco.getRange(Linha,113).setValue(Entrada.getRange('D40').getValue());
    SheetBanco.getRange(Linha,114).setValue(Entrada.getRange('D41').getValue());
    SheetBanco.getRange(Linha,115).setValue(Entrada.getRange('D42').getValue());
    SheetBanco.getRange(Linha,116).setValue(Entrada.getRange('D43').getValue());
    SheetBanco.getRange(Linha,117).setValue(Entrada.getRange('D44').getValue());
    SheetBanco.getRange(Linha,118).setValue(Entrada.getRange('D45').getValue());
    SheetBanco.getRange(Linha,119).setValue(Entrada.getRange('D46').getValue());
    SheetBanco.getRange(Linha,120).setValue(Entrada.getRange('D47').getValue());
    SheetBanco.getRange(Linha,121).setValue(Entrada.getRange('D48').getValue());
    SheetBanco.getRange(Linha,122).setValue(Entrada.getRange('D49').getValue());
    SheetBanco.getRange(Linha,123).setValue(Entrada.getRange('D50').getValue());
    SheetBanco.getRange(Linha,124).setValue(Entrada.getRange('D51').getValue());
    SheetBanco.getRange(Linha,125).setValue(Entrada.getRange('D52').getValue());
    SheetBanco.getRange(Linha,126).setValue(Entrada.getRange('D53').getValue());
    SheetBanco.getRange(Linha,127).setValue(Entrada.getRange('D54').getValue());
    SheetBanco.getRange(Linha,128).setValue(Entrada.getRange('F39').getValue());
    SheetBanco.getRange(Linha,129).setValue(Entrada.getRange('F40').getValue());
    SheetBanco.getRange(Linha,130).setValue(Entrada.getRange('F41').getValue());
    SheetBanco.getRange(Linha,131).setValue(Entrada.getRange('F42').getValue());
    SheetBanco.getRange(Linha,132).setValue(Entrada.getRange('F43').getValue());
    SheetBanco.getRange(Linha,133).setValue(Entrada.getRange('F44').getValue());
    SheetBanco.getRange(Linha,134).setValue(Entrada.getRange('F45').getValue());
    SheetBanco.getRange(Linha,135).setValue(Entrada.getRange('F46').getValue());
    SheetBanco.getRange(Linha,136).setValue(Entrada.getRange('F47').getValue());
    SheetBanco.getRange(Linha,137).setValue(Entrada.getRange('F48').getValue());
    SheetBanco.getRange(Linha,138).setValue(Entrada.getRange('F49').getValue());
    SheetBanco.getRange(Linha,139).setValue(Entrada.getRange('F50').getValue());
    SheetBanco.getRange(Linha,140).setValue(Entrada.getRange('F51').getValue());
    SheetBanco.getRange(Linha,141).setValue(Entrada.getRange('F52').getValue());
    SheetBanco.getRange(Linha,142).setValue(Entrada.getRange('F53').getValue());
    SheetBanco.getRange(Linha,143).setValue(Entrada.getRange('F54').getValue());
    SheetBanco.getRange(Linha,144).setValue(Entrada.getRange('H39').getValue());
    SheetBanco.getRange(Linha,145).setValue(Entrada.getRange('H40').getValue());
    SheetBanco.getRange(Linha,146).setValue(Entrada.getRange('H41').getValue());
    SheetBanco.getRange(Linha,147).setValue(Entrada.getRange('H42').getValue());
    SheetBanco.getRange(Linha,148).setValue(Entrada.getRange('H43').getValue());
    SheetBanco.getRange(Linha,149).setValue(Entrada.getRange('H44').getValue());
    SheetBanco.getRange(Linha,150).setValue(Entrada.getRange('H45').getValue());
    SheetBanco.getRange(Linha,151).setValue(Entrada.getRange('H46').getValue());
    SheetBanco.getRange(Linha,152).setValue(Entrada.getRange('H47').getValue());
    SheetBanco.getRange(Linha,153).setValue(Entrada.getRange('H48').getValue());
    SheetBanco.getRange(Linha,154).setValue(Entrada.getRange('H49').getValue());
    SheetBanco.getRange(Linha,155).setValue(Entrada.getRange('H50').getValue());
    SheetBanco.getRange(Linha,156).setValue(Entrada.getRange('H51').getValue());
    SheetBanco.getRange(Linha,157).setValue(Entrada.getRange('H52').getValue());
    SheetBanco.getRange(Linha,158).setValue(Entrada.getRange('H53').getValue());
    SheetBanco.getRange(Linha,159).setValue(Entrada.getRange('H54').getValue());

    SheetBanco.getRange(Linha,160).setValue(Entrada.getRange('F24').getValue());
    SheetBanco.getRange(Linha,161).setValue(Entrada.getRange('F25').getValue());
    SheetBanco.getRange(Linha,162).setValue(Entrada.getRange('F26').getValue());
    SheetBanco.getRange(Linha,163).setValue(Entrada.getRange('F27').getValue());
    SheetBanco.getRange(Linha,164).setValue(Entrada.getRange('F28').getValue());
    SheetBanco.getRange(Linha,165).setValue(Entrada.getRange('F36').getValue());
    
    SheetBanco.getRange(Linha,166).setValue(Entrada.getRange('D56').getValue());
    SheetBanco.getRange(Linha,167).setValue(Entrada.getRange('D57').getValue());
    SheetBanco.getRange(Linha,168).setValue(Entrada.getRange('D58').getValue());
    SheetBanco.getRange(Linha,169).setValue(Entrada.getRange('D55').getValue());

    SheetBanco.getRange(Linha,170).setValue(Entrada.getRange('H56').getValue());
    SheetBanco.getRange(Linha,171).setValue(Entrada.getRange('H57').getValue());
    SheetBanco.getRange(Linha,172).setValue(Entrada.getRange('H58').getValue());
    SheetBanco.getRange(Linha,173).setValue(Entrada.getRange('H55').getValue());
     
    SheetBanco.getRange(Linha,174).setValue(Entrada.getRange('H2').getValue());
    SheetBanco.getRange(Linha,175).setValue(Entrada.getRange('H3').getValue());
    SheetBanco.getRange(Linha,176).setValue(Entrada.getRange('H4').getValue());
    
    SheetBanco.getRange(Linha,177).setValue(Entrada.getRange('H21').getValue());
  
    SheetBanco.getRange(Linha,178).setValue(Entrada.getRange('B55').getValue());
  
    SheetBanco.getRange(Linha,179).setValue(Entrada.getRange('F55').getValue());
  
    SheetBanco.getRange(Linha,180).setValue(Entrada.getRange('F22').getValue());
  
    LimparEntrada();
    
  
};

function ExcluirOcorrencia(){
  
  var Entrada = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Entrada'); 
  var Ocorrencia = Entrada.getRange('B2').getValue();
  
 
  var SheetBanco = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Banco');
    
  SheetBanco.getActiveCell();
  
  var LocalPesquisa = SheetBanco.getRange(2,2, SheetBanco.getLastRow()).getValues();
  var Resultado = LocalPesquisa.Pesquisa(Ocorrencia);
  
  var Linha1 = Resultado + 2
  
  
  if(Resultado != -1){
    
    LocalizarCelulaVazia_Excluida()
  
    var Excluida = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Excluidas');
    var Linha2 = Excluida.getCurrentCell().getRow();
       
      
    Excluida.getRange(Linha2, 2).setValue(SheetBanco.getRange(Linha1, 2).getValue());
    Excluida.getRange(Linha2, 3).setValue(SheetBanco.getRange(Linha1, 3).getValue());
    Excluida.getRange(Linha2, 4).setValue(SheetBanco.getRange(Linha1, 5).getValue());
    Excluida.getRange(Linha2, 5).setValue(SheetBanco.getRange(Linha1, 6).getValue());
    Excluida.getRange(Linha2, 6).setValue(SheetBanco.getRange(Linha1, 11).getValue());
    Excluida.getRange(Linha2,7).setValue(SheetBanco.getRange(Linha1, 13).getValue());
    Excluida.getRange(Linha2,8).setValue(SheetBanco.getRange(Linha1, 22).getValue());
    Excluida.getRange(Linha2,9).setValue(SheetBanco.getRange(Linha1, 29).getValue());
    Excluida.getRange(Linha2,10).setValue(SheetBanco.getRange(Linha1, 41).getValue());
    Excluida.getRange(Linha2,11).setValue(SheetBanco.getRange(Linha1, 42).getValue());
    Excluida.getRange(Linha2,12).setValue(SheetBanco.getRange(Linha1, 45).getValue());
    Excluida.getRange(Linha2,13).setValue(SheetBanco.getRange(Linha1, 48).getValue());
    Excluida.getRange(Linha2,14).setValue(SheetBanco.getRange(Linha1, 66).getValue());
    Excluida.getRange(Linha2,15).setValue(SheetBanco.getRange(Linha1, 69).getValue());
    Excluida.getRange(Linha2,16).setValue(SheetBanco.getRange(Linha1, 72).getValue());
    Excluida.getRange(Linha2,17).setValue(SheetBanco.getRange(Linha1, 80).getValue());
    Excluida.getRange(Linha2,18).setValue(SheetBanco.getRange(Linha1, 40).getValue());
        
        
    SheetBanco.deleteRow(Linha1);
    
     
    Browser.msgBox("Ocorrência excluída!")
  
    LimparEntrada();
  }else{
    
    Browser.msgBox("Ocorrência não localizada!")
  
  }
  
};

Array.prototype.Pesquisa = function(Procura){
  if(Procura == "") return false;
  for(var Linha = 0; Linha < this.length;Linha++)
    if(this[Linha] == Procura) return Linha; 
   return -1

}




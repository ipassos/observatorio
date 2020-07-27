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




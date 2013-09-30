//UEwasp.h

//************************************************************
//     作    者：                                            *
//                邝祝芳                                     *
//     文件名称：                                            * 
//               DLL外部接口(水和蒸汽性质计算动态连接库)     *
//     完成时间：                                            *
//                2005年5月                                  *
//************************************************************

extern "C" _declspec(dllimport) void SETSTD_WASP(int  STDID);
extern "C" _declspec(dllimport) void GETSTD_WASP(int  STDID);
extern "C" _declspec(dllimport) void P2T(double P, double * T, int * RANGE );
extern "C" _declspec(dllimport) void P2HL(double P, double * H , int * RANGE);
extern "C" _declspec(dllimport) void P2HG(double P  ,double * H , int * RANGE);
extern "C" _declspec(dllimport) void P2SL(double P  ,double * S , int * RANGE);
extern "C" _declspec(dllimport) void P2SG(double P  ,double * S , int * RANGE);
extern "C" _declspec(dllimport) void P2VL(double P  ,double * V , int * RANGE);
extern "C" _declspec(dllimport) void P2VG(double P  ,double * V , int * RANGE);
extern "C" _declspec(dllimport) void P2L(double P   ,double * T,double * H, double *S, double *V, double *X, int * RANGE);
extern "C" _declspec(dllimport) void P2G(double P   ,double * T,double * H, double *S, double *V, double *X, int * RANGE);
extern "C" _declspec(dllimport) void P2CPL(double P ,double *CP,int * RANGE);
extern "C" _declspec(dllimport) void P2CPG(double P ,double *CP,int * RANGE);
extern "C" _declspec(dllimport) void P2CVL(double P ,double *CV,int * RANGE);
extern "C" _declspec(dllimport) void P2CVG(double P ,double *CV,int * RANGE);
extern "C" _declspec(dllimport) void P2EL(double P  ,double * E,int    * RANGE);
extern "C" _declspec(dllimport) void P2EG(double P  ,double * E,int    * RANGE);
extern "C" _declspec(dllimport) void P2SSPL(double P,double *SSP,int   * RANGE);
extern "C" _declspec(dllimport) void P2SSPG(double P,double *SSP,int   * RANGE);
extern "C" _declspec(dllimport) void P2KSL(double P ,double *KS ,int   * RANGE);
extern "C" _declspec(dllimport) void P2KSG(double P ,double *KS ,int   * RANGE);
extern "C" _declspec(dllimport) void P2ETAL(double P,double *ETA,int * RANGE);
extern "C" _declspec(dllimport) void P2ETAG(double P,double *ETA,int * RANGE);
extern "C" _declspec(dllimport) void P2UL(double P  ,double *U,int * RANGE);
extern "C" _declspec(dllimport) void P2UG(double P  ,double *U,int * RANGE);
extern "C" _declspec(dllimport) void P2RAMDL(double P ,double *RAMD, int * RANGE);
extern "C" _declspec(dllimport) void P2RAMDG(double P ,double *RAMD, int * RANGE);
extern "C" _declspec(dllimport) void P2PRNL(double P  ,double *PRN,  int * RANGE);
extern "C" _declspec(dllimport) void P2PRNG(double P  ,double *PRN,  int * RANGE);
extern "C" _declspec(dllimport) void P2EPSL(double P  ,double *EPS,  int * RANGE);
extern "C" _declspec(dllimport) void P2EPSG(double P  ,double *EPS,  int * RANGE);
extern "C" _declspec(dllimport) void P2NL(double P, double Lamd,double *N,int * RANGE);
extern "C" _declspec(dllimport) void P2NG(double P, double Lamd,double *N,int * RANGE);

extern "C" _declspec(dllimport) void PT2H(double P, double T,double *H,int * RANGE);
extern "C" _declspec(dllimport) void PT2S(double P, double T,double *S,int * RANGE);
extern "C" _declspec(dllimport) void PT2V(double P, double T,double *V,int * RANGE);
extern "C" _declspec(dllimport) void PT2X(double P, double T,double *X,int * RANGE);
extern "C" _declspec(dllimport) void PT(double P, double T, double * H, double *S, double *V, double *X, int * RANGE);
extern "C" _declspec(dllimport) void PT2CP(double P, double T,double *CP,int * RANGE);
extern "C" _declspec(dllimport) void PT2CV(double P, double T,double *CV,int * RANGE);
extern "C" _declspec(dllimport) void PT2E(double P, double T,double *E,int * RANGE);
extern "C" _declspec(dllimport) void PT2SSP(double P, double T,double *SSP,int * RANGE);
extern "C" _declspec(dllimport) void PT2KS(double P, double T,double *KS,int * RANGE);
extern "C" _declspec(dllimport) void PT2ETA(double P, double T,double *ETA,int * RANGE);
extern "C" _declspec(dllimport) void PT2U(double P, double T,double *U,int * RANGE);
extern "C" _declspec(dllimport) void PT2RAMD(double P, double T,double *RAMD,int * RANGE);
extern "C" _declspec(dllimport) void PT2PRN(double P, double T,double *PRN,int * RANGE);
extern "C" _declspec(dllimport) void PT2EPS(double P, double T,double *EPS,int * RANGE);
extern "C" _declspec(dllimport) void PT2N(double P, double T,double LAMD, double *N,int * RANGE);

extern "C" _declspec(dllimport) void PH2T(double P, double H,double *T,int * RANGE);
extern "C" _declspec(dllimport) void PH2S(double P, double H,double *S,int * RANGE);
extern "C" _declspec(dllimport) void PH2V(double P, double H,double *V,int * RANGE);
extern "C" _declspec(dllimport) void PH2X(double P, double H,double *X,int * RANGE);
extern "C" _declspec(dllimport) void PH(double P  ,double *T,double H, double * S,double *V,double *X,int * RANGE);

extern "C" _declspec(dllimport) void PS2T(double P, double S,double *T,int * RANGE);
extern "C" _declspec(dllimport) void PS2H(double P, double S,double *H,int * RANGE);
extern "C" _declspec(dllimport) void PS2V(double P, double S,double *V,int * RANGE);
extern "C" _declspec(dllimport) void PS2X(double P, double S,double *X,int * RANGE);
extern "C" _declspec(dllimport) void PS(double P  ,double *T, double *H,double S,double *V, double *X,int * RANGE);

extern "C" _declspec(dllimport) void PV2T(double P, double V,double *T,int * RANGE);
extern "C" _declspec(dllimport) void PV2H(double P, double V,double *H,int * RANGE);
extern "C" _declspec(dllimport) void PV2S(double P, double V,double *S,int * RANGE);
extern "C" _declspec(dllimport) void PV2X(double P, double V,double *X,int * RANGE);
extern "C" _declspec(dllimport) void PV(double P  ,double *T, double *H, double *S,double V,double *X,int * RANGE);

extern "C" _declspec(dllimport) void PX2T(double P, double X,double *T,int * RANGE);
extern "C" _declspec(dllimport) void PX2H(double P, double X,double *H,int * RANGE);
extern "C" _declspec(dllimport) void PX2S(double P, double X,double *S,int * RANGE);
extern "C" _declspec(dllimport) void PX2V(double P, double X,double *V,int * RANGE);
extern "C" _declspec(dllimport) void PX(double P  ,double *T, double *H, double *S, double *V,double X,int * RANGE);

extern "C" _declspec(dllimport) void T2P(double T  ,double *P, int * RANGE);
extern "C" _declspec(dllimport) void T2HL(double T  ,double *H, int * RANGE);
extern "C" _declspec(dllimport) void T2HG(double T  ,double *H, int * RANGE);
extern "C" _declspec(dllimport) void T2SL(double T  ,double *S, int * RANGE);
extern "C" _declspec(dllimport) void T2SG(double T  ,double *S, int * RANGE);
extern "C" _declspec(dllimport) void T2VL(double T  ,double *V, int * RANGE);
extern "C" _declspec(dllimport) void T2VG(double T  ,double *V, int * RANGE);
extern "C" _declspec(dllimport) void T2L(double *P,double T  ,double *H,double *S,double *V,double *X,int * RANGE);
extern "C" _declspec(dllimport) void T2G(double *P,double T  ,double *H,double *S,double *V,double *X,int * RANGE);
extern "C" _declspec(dllimport) void T2CPL(double T  ,double *CP,int * RANGE);
extern "C" _declspec(dllimport) void T2CPG(double T  ,double *CP,int * RANGE);
extern "C" _declspec(dllimport) void T2CVL(double T  ,double *CV,int * RANGE);
extern "C" _declspec(dllimport) void T2CVG(double T  ,double *CV,int * RANGE);
extern "C" _declspec(dllimport) void T2EL(double T  ,double *E,int * RANGE);
extern "C" _declspec(dllimport) void T2EG(double T  ,double *E,int * RANGE);
extern "C" _declspec(dllimport) void T2SSPL(double T  ,double *SSP,int * RANGE);
extern "C" _declspec(dllimport) void T2SSPG(double T  ,double *SSP,int * RANGE);
extern "C" _declspec(dllimport) void T2KSL(double T  ,double *KS,int * RANGE);
extern "C" _declspec(dllimport) void T2KSG(double T  ,double *KS,int * RANGE);
extern "C" _declspec(dllimport) void T2ETAL(double T  ,double *ETA,int * RANGE);
extern "C" _declspec(dllimport) void T2ETAG(double T  ,double *ETA,int * RANGE);
extern "C" _declspec(dllimport) void T2UL(double T  ,double *U,int * RANGE);
extern "C" _declspec(dllimport) void T2UG(double T  ,double *U,int * RANGE);
extern "C" _declspec(dllimport) void T2RAMDL(double T  ,double *RAMD,int * RANGE);
extern "C" _declspec(dllimport) void T2RAMDG(double T  ,double *RAMD,int * RANGE);
extern "C" _declspec(dllimport) void T2PRNL(double T  ,double *PRN,int * RANGE);
extern "C" _declspec(dllimport) void T2PRNG(double T  ,double *PRN,int * RANGE);
extern "C" _declspec(dllimport) void T2EPSL(double T  ,double *EPS,int * RANGE);
extern "C" _declspec(dllimport) void T2EPSG(double T  ,double *EPS,int * RANGE);
extern "C" _declspec(dllimport) void T2NL(double T, double Lamd,double *N,int * RANGE);
extern "C" _declspec(dllimport) void T2NG(double T, double Lamd,double *N,int * RANGE);
extern "C" _declspec(dllimport) void T2SURFT(double T  ,double *SURFT,int * RANGE);

extern "C" _declspec(dllimport) void TH2P(double T, double H,double *P,int * RANGE);
extern "C" _declspec(dllimport) void TH2PLP(double T, double H,double *P,int * RANGE);
extern "C" _declspec(dllimport) void TH2PHP(double T, double H,double *P,int * RANGE);
extern "C" _declspec(dllimport) void TH2S(double T, double H,double *S,int * RANGE);
extern "C" _declspec(dllimport) void TH2SLP(double T, double H,double *S,int * RANGE);
extern "C" _declspec(dllimport) void TH2SHP(double T, double H,double *S,int * RANGE);
extern "C" _declspec(dllimport) void TH2V(double T, double H,double *V,int * RANGE);
extern "C" _declspec(dllimport) void TH2VLP(double T, double H,double *V,int * RANGE);
extern "C" _declspec(dllimport) void TH2VHP(double T, double H,double *V,int * RANGE);
extern "C" _declspec(dllimport) void TH2X(double T, double H,double *X,int * RANGE);
extern "C" _declspec(dllimport) void TH2XLP(double T, double H,double *X,int * RANGE);
extern "C" _declspec(dllimport) void TH2XHP(double T, double H,double *X,int * RANGE);
extern "C" _declspec(dllimport) void TH(double *P,double T, double H,double * S,double *V,double *X,int * RANGE);
extern "C" _declspec(dllimport) void THLP(double *P,double T, double H,double * S,double *V,double *X,int * RANGE);
extern "C" _declspec(dllimport) void THHP(double *P,double T, double H,double * S,double *V,double *X,int * RANGE);

extern "C" _declspec(dllimport) void TS2PHP(double T, double S,double *P,int * RANGE);
extern "C" _declspec(dllimport) void TS2PLP(double T, double S,double *P,int * RANGE);
extern "C" _declspec(dllimport) void TS2P(double T, double S,double *P,int * RANGE);
extern "C" _declspec(dllimport) void TS2HHP(double T, double S,double *H,int * RANGE);
extern "C" _declspec(dllimport) void TS2HLP(double T, double S,double *H,int * RANGE);
extern "C" _declspec(dllimport) void TS2H(double T, double S,double *H,int * RANGE);
extern "C" _declspec(dllimport) void TS2VHP(double T, double S,double *V,int * RANGE);
extern "C" _declspec(dllimport) void TS2VLP(double T, double S,double *V,int * RANGE);
extern "C" _declspec(dllimport) void TS2V(double T, double S,double *V,int * RANGE);
extern "C" _declspec(dllimport) void TS2X(double T, double S,double *X,int * RANGE);
extern "C" _declspec(dllimport) void TSHP(double *P,double T,double *H,double S,double *V, double *X,int * RANGE);
extern "C" _declspec(dllimport) void TSLP(double *P,double T,double *H,double S,double *V, double *X,int * RANGE);
extern "C" _declspec(dllimport) void TS(double *P,double T, double *H,double S,double *V, double *X,int * RANGE);

extern "C" _declspec(dllimport) void TV2P(double T, double V,double *P,int * RANGE);
extern "C" _declspec(dllimport) void TV2H(double T, double V,double *H,int * RANGE);
extern "C" _declspec(dllimport) void TV2S(double T, double V,double *S,int * RANGE);
extern "C" _declspec(dllimport) void TV2X(double T, double V,double *X,int * RANGE);
extern "C" _declspec(dllimport) void TV(double *P,double T  ,double *H,double *S ,double V , double *X,int * RANGE);

extern "C" _declspec(dllimport) void TX2P(double T, double X,double *P,int * RANGE);
extern "C" _declspec(dllimport) void TX2H(double T, double X,double *H,int * RANGE);
extern "C" _declspec(dllimport) void TX2S(double T, double X,double *S,int * RANGE);
extern "C" _declspec(dllimport) void TX2V(double T, double X,double *V,int * RANGE);
extern "C" _declspec(dllimport) void TX(double *P,double T  ,double *H, double *S, double *V,double X,int * RANGE);

extern "C" _declspec(dllimport) void H2TL(double H, double *T,int * RANGE);
extern "C" _declspec(dllimport) void HS2P(double H, double S,double *P,int * RANGE);
extern "C" _declspec(dllimport) void HS2T(double H, double S,double *T,int * RANGE);
extern "C" _declspec(dllimport) void HS2V(double H, double S,double *V,int * RANGE);
extern "C" _declspec(dllimport) void HS2X(double H, double S,double *X,int * RANGE);
extern "C" _declspec(dllimport) void HS(double *P, double *T,double H, double S,double *V, double *X,int * RANGE);

extern "C" _declspec(dllimport) void HV2P(double H, double V,double *P,int * RANGE);
extern "C" _declspec(dllimport) void HV2T(double H, double V,double *T,int * RANGE);
extern "C" _declspec(dllimport) void HV2S(double H, double V,double *S,int * RANGE);
extern "C" _declspec(dllimport) void HV2X(double H, double V,double *X,int * RANGE);
extern "C" _declspec(dllimport) void HV(double *P, double *T,double H,double *S,double V,double *X,int * RANGE);

extern "C" _declspec(dllimport) void HX2P(double H, double X,double *P,int * RANGE);
extern "C" _declspec(dllimport) void HX2PLP(double H, double X,double *P,int * RANGE);
extern "C" _declspec(dllimport) void HX2PHP(double H, double X,double *P,int * RANGE);
extern "C" _declspec(dllimport) void HX2T(double H, double X,double *T,int * RANGE);
extern "C" _declspec(dllimport) void HX2TLP(double H, double X,double *T,int * RANGE);
extern "C" _declspec(dllimport) void HX2THP(double H, double X,double *T,int * RANGE);
extern "C" _declspec(dllimport) void HX2S(double H, double X,double *S,int * RANGE);
extern "C" _declspec(dllimport) void HX2SLP(double H, double X,double *S,int * RANGE);
extern "C" _declspec(dllimport) void HX2SHP(double H, double X,double *S,int * RANGE);
extern "C" _declspec(dllimport) void HX2V(double H, double X,double *V,int * RANGE);
extern "C" _declspec(dllimport) void HX2VLP(double H, double X,double *V,int * RANGE);
extern "C" _declspec(dllimport) void HX2VHP(double H, double X,double *V,int * RANGE);
extern "C" _declspec(dllimport) void HX(double *P, double *T,double H,double *S, double *V,double X,int * RANGE);
extern "C" _declspec(dllimport) void HXLP(double *P, double *T,double H,double *S, double *V,double X,int * RANGE);
extern "C" _declspec(dllimport) void HXHP(double *P, double *T,double H,double *S, double *V,double X,int * RANGE);

extern "C" _declspec(dllimport) void S2TG(double S,double *T,int * RANGE);
extern "C" _declspec(dllimport) void SV2P(double S, double V,double *P,int * RANGE);
extern "C" _declspec(dllimport) void SV2T(double S, double V,double *T,int * RANGE);
extern "C" _declspec(dllimport) void SV2H(double S, double V,double *H,int * RANGE);
extern "C" _declspec(dllimport) void SV2X(double S, double V,double *X,int * RANGE);
extern "C" _declspec(dllimport) void SV(double *P, double *T, double *H,double S, double V,double *X,int * RANGE);

extern "C" _declspec(dllimport) void SX2P(double S, double X,double *P,int * RANGE);
extern "C" _declspec(dllimport) void SX2PLP(double S, double X,double *P,int * RANGE);
extern "C" _declspec(dllimport) void SX2PMP(double S, double X,double *P,int * RANGE);
extern "C" _declspec(dllimport) void SX2PHP(double S, double X,double *P,int * RANGE);
extern "C" _declspec(dllimport) void SX2TLP(double S, double X,double *T,int * RANGE);
extern "C" _declspec(dllimport) void SX2TMP(double S, double X,double *T,int * RANGE);
extern "C" _declspec(dllimport) void SX2THP(double S, double X,double *T,int * RANGE);
extern "C" _declspec(dllimport) void SX2T(double S, double X,double *T,int * RANGE);
extern "C" _declspec(dllimport) void SX2H(double S, double X,double *H,int * RANGE);
extern "C" _declspec(dllimport) void SX2HLP(double S, double X,double *H,int * RANGE);
extern "C" _declspec(dllimport) void SX2HMP(double S, double X,double *H,int * RANGE);
extern "C" _declspec(dllimport) void SX2HHP(double S, double X,double *H,int * RANGE);
extern "C" _declspec(dllimport) void SX2V(double S, double X,double *V,int * RANGE);
extern "C" _declspec(dllimport) void SX2VLP(double S, double X,double *V,int * RANGE);
extern "C" _declspec(dllimport) void SX2VMP(double S, double X,double *V,int * RANGE);
extern "C" _declspec(dllimport) void SX2VHP(double S, double X,double *V,int * RANGE);
extern "C" _declspec(dllimport) void SX(double *P, double *T, double *H,double S,double *V,double X,int * RANGE);
extern "C" _declspec(dllimport) void SXLP(double *P, double *T, double *H,double S,double *V,double X,int * RANGE);
extern "C" _declspec(dllimport) void SXMP(double *P, double *T, double *H,double S,double *V,double X,int * RANGE);
extern "C" _declspec(dllimport) void SXHP(double *P, double *T, double *H,double S,double *V,double X,int * RANGE);

extern "C" _declspec(dllimport) void V2TG(double V , double *T,int * RANGE);
extern "C" _declspec(dllimport) void VX2P(double V, double X,double *P,int * RANGE);
extern "C" _declspec(dllimport) void VX2PLP(double V, double X,double *P,int * RANGE);
extern "C" _declspec(dllimport) void VX2PHP(double V, double X,double *P,int * RANGE);
extern "C" _declspec(dllimport) void VX2T(double V, double X,double *T,int * RANGE);
extern "C" _declspec(dllimport) void VX2TLP(double V, double X,double *T,int * RANGE);
extern "C" _declspec(dllimport) void VX2THP(double V, double X,double *T,int * RANGE);
extern "C" _declspec(dllimport) void VX2H(double V, double X,double *H,int * RANGE);
extern "C" _declspec(dllimport) void VX2HLP(double V, double X,double *H,int * RANGE);
extern "C" _declspec(dllimport) void VX2HHP(double V, double X,double *H,int * RANGE);
extern "C" _declspec(dllimport) void VX2S(double V, double X,double *S,int * RANGE);
extern "C" _declspec(dllimport) void VX2SLP(double V, double X,double *S,int * RANGE);
extern "C" _declspec(dllimport) void VX2SHP(double V, double X,double *S,int * RANGE);
extern "C" _declspec(dllimport) void VX(double *P, double *T, double *H, double *S,double V, double X,int * RANGE);
extern "C" _declspec(dllimport) void VXLP(double *P, double *T, double *H, double *S,double V, double X,int * RANGE);
extern "C" _declspec(dllimport) void VXHP(double *P, double *T, double *H, double *S,double V, double X,int * RANGE);
'// Version 3.4.10                          ' requires >= Excel 2000 (for long lines [for Array assignment])
'// Thanks to Jerry W. Lewis for lots of help with testing and improvements to the code.

'// Copyright © [2022] [Ian Smith]

'// Permission is hereby granted, free of charge, to any person obtaining a copy
'// of this software and associated documentation files (the "Software"), to deal
'// in the Software without restriction, including without limitation the rights
'// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'// copies of the Software, and to permit persons to whom the Software is
'// furnished to do so, subject to the following conditions:

'// The above copyright notice and this permission notice shall be included in all
'// copies or substantial portions of the Software.

'// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
'// SOFTWARE.

Option Explicit

Type TValues
   value As Double
   Log2Adds  As Integer
End Type

Type TAddStack
   Store As Boolean
   Where As Integer
   Stack(50) As TValues
End Type

Const NonIntegralValuesAllowed_df = True    ' Are non-integral degrees of freedom for t, chi_square and f distributions allowed?
Const NonIntegralValuesAllowed_NB = True    ' Is "successes required" parameter for negative binomial allowed to be non-integral?

Const NonIntegralValuesAllowed_Others = True ' Can Int function be applied to parameters like sample_size or is it a fault if the parameter is non-integral?

Const nc_limit = 1000000#                   ' Upper Limit for non-centrality parameters - as far as I know it's ok but slower and slower up to 1e12. Above that I don't know.
Const sumAcc = 5e-16
Const cfSmall = 0.00000000000001
Const cfVSmall = 0.000000000000001
Const minLog1Value = -0.79149064
Const OneOverSqrTwoPi = 0.39894228040143267793994605993438 ' 0.39894228040143267793994605993438
Const scalefactor = 1.1579208923731619542357098500869e+77  ' 1.1579208923731619542357098500869e+77 = 2^256  ' used for rescaling calcs w/o impacting accuracy, to avoid over/underflow
Const scalefactor2 = 8.6361685550944446253863518628004e-78 ' 8.6361685550944446253863518628004e-78 = 2^-256
Const max_discrete = 9007199254740991                      ' 2^53 required for exact addition of 1 in hypergeometric routines
Const max_crit = 4503599627370496                          ' 2^52 to make sure plenty of room for exact addition of 1 in crit routines
Const nearly_zero = 9.99999983659714E-317
Const cSmall = 5.562684646268003457725581793331e-309       ' (smallest number before we start losing precision)/4
Const excel0 = 2.2250738585E-308                           ' (number smaller than those excel replaces with 0)
Const t_nc_limit = 1.3407807929942597e+154                 ' just under 1/Sqr(cSmall)
Const Log1p5 = 0.40546510810816438197801311546435          ' 0.40546510810816438197801311546435 = Log(1.5)
Const logfbit0p5 = 0.054814121051917653896138702348386     ' 0.054814121051917653896138702348386 = logfbit(0.5)
Const lfb_0 = 8.1061466795327258219670263594382e-02        ' 8.1061466795327258219670263594382e-02 = logfbit(0#)
Const lfb_1 = 4.1340695955409294093822081407118e-02        ' 4.1340695955409294093822081407118e-02 = logfbit(1#)
Const lfb_2 = 2.7677925684998339148789292746245e-02        ' 2.7677925684998339148789292746245e-02 = logfbit(2#)
Const lfb_3 = 2.0790672103765093111522771767849e-02        ' 2.0790672103765093111522771767849e-02 = logfbit(3#)
Const lfb_4 = 1.6644691189821192163194865373593e-02        ' 1.6644691189821192163194865373593e-02 = logfbit(4#)
Const lfb_5 = 1.3876128823070747998745727023763e-02        ' 1.3876128823070747998745727023763e-02 = logfbit(5#)

'For logfbit functions                       ' Stirling's series for ln(Gamma(x)), A046968/A046969
Const lfbc1 = 1# / 12#
Const lfbc2 = 1# / 30#                       ' lfbc2 on are Sloane's ratio times 12
Const lfbc3 = 1# / 105#
Const lfbc4 = 1# / 140#
Const lfbc5 = 1# / 99#
Const lfbc6 = 691# / 30030#
Const lfbc7 = 1# / 13#
Const lfbc8 = .35068485511628418514 '.35068485511628418514   ' Chosen to make logfbit(6) & logfbit(7) correct
Const lfbc9 = 1.6769380337122674863 '1.6769380337122674863   ' Chosen to make logfbit(6) & logfbit(7) correct

'For logfbit functions                      'Stieltjes' continued fraction
Const cf_0 = 1# / 12#
Const cf_1 = 1# / 30#
Const cf_2 = 53# / 210#
Const cf_3 = 195# / 371#
Const cf_4 = 22999# / 22737#
Const cf_5 = 29944523# / 19733142#
Const cf_6 = 109535241009# / 48264275462#
Const cf_7 = 3.0099173832593981700731407342077  '3.0099173832593981700731407342077
Const cf_8 = 4.026887192343901226168879531814   '4.026887192343901226168879531814
Const cf_9 = 5.0027680807540300516885024122767  '5.0027680807540300516885024122767
Const cf_10 = 6.2839113708157821800726631549524 '6.2839113708157821800726631549524
Const cf_11 = 7.4959191223840339297523547082674 '7.4959191223840339297523547082674
Const cf_12 = 9.0406602343677266995311393604326 '9.0406602343677266995311393604326
Const cf_13 = 10.489303654509482277188371304593 '10.489303654509482277188371304593
Const cf_14 = 12.297193610386205863989437140092 '12.297193610386205863989437140092
Const cf_15 = 13.982876953992430188259760651279 '13.982876953992430188259760651279
Const cf_16 = 16.053551416704935469715616365007 '16.053551416704935469715616365007
Const cf_17 = 17.976607399870277592569472307671 '17.976607399870277592569472307671
Const cf_18 = 20.309762027441653743805414720495 '20.309762027441653743805414720495
Const cf_19 = 22.470471639933132495517941571508 '22.470471639933132495517941571508
Const cf_20 = 25.065846548945972029163400322506 '25.065846548945972029163400322506
Const cf_21 = 27.464451825029133609175558982646 '27.464451825029133609175558982646
Const cf_22 = 30.321821231673047126882599306406 '30.321821231673047126882599306406
Const cf_23 = 32.958533929972987219994066451412 '32.958533929972987219994066451412
Const cf_24 = 36.077698931299242645153320900855 '36.077698931299242645153320900855
Const cf_25 = 38.952706682311555734544390410481 '38.952706682311555734544390410481
Const cf_26 = 42.333490043576957211381853948856 '42.333490043576957211381853948856
Const cf_27 = 45.446960850061621014424175737541 '45.446960850061621014424175737541
Const cf_28 = 49.089203129012597708164883350275 '49.089203129012597708164883350275
Const cf_29 = 52.441288751415337312569856046996 '52.441288751415337312569856046996
Const cf_30 = 56.344845345341843538420365947476 '56.344845345341843538420365947476
Const cf_31 = 59.935683907165858207852583492752 '59.935683907165858207852583492752
Const cf_32 = 64.100422755920354527906611892238 '64.100422755920354527906611892238
Const cf_33 = 67.930140788018221186367702745199 '67.930140788018221186367702745199
Const cf_34 = 72.355940555211701969680052963236 '72.355940555211701969680052963236
Const cf_35 = 76.424654626829689752585090422288 '76.424654626829689752585090422288
Const cf_36 = 81.111403237247965484814230985683 '81.111403237247965484814230985683
Const cf_37 = 85.419221276410972614585638717349 '85.419221276410972614585638717349
Const cf_38 = 90.366814723864108595513574581683 '90.366814723864108595513574581683
Const cf_39 = 94.913837100009887953076231291987 '94.913837100009887953076231291987
Const cf_40 = 100.12217846392919748899074683447 '100.12217846392919748899074683447


'For invcnormal                             ' http://lib.stat.cmu.edu/apstat/241
Const a0 = 3.3871328727963666080            ' 3.3871328727963666080
Const a1 = 133.14166789178437745            ' 133.14166789178437745
Const a2 = 1971.5909503065514427            ' 1971.5909503065514427
Const a3 = 13731.693765509461125            ' 13731.693765509461125
Const a4 = 45921.953931549871457            ' 45921.953931549871457
Const a5 = 67265.770927008700853            ' 67265.770927008700853
Const a6 = 33430.575583588128105            ' 33430.575583588128105
Const a7 = 2509.0809287301226727            ' 2509.0809287301226727
Const b1 = 42.313330701600911252            ' 42.313330701600911252
Const b2 = 687.18700749205790830            ' 687.18700749205790830
Const b3 = 5394.1960214247511077            ' 5394.1960214247511077
Const b4 = 21213.794301586595867            ' 21213.794301586595867
Const b5 = 39307.895800092710610            ' 39307.895800092710610
Const b6 = 28729.085735721942674            ' 28729.085735721942674
Const b7 = 5226.4952788528545610            ' 5226.4952788528545610
'//Coefficients for P not close to 0, 0.5 or 1.
Const c0 = 1.42343711074968357734           ' 1.42343711074968357734
Const c1 = 4.63033784615654529590           ' 4.63033784615654529590
Const c2 = 5.76949722146069140550           ' 5.76949722146069140550
Const c3 = 3.64784832476320460504           ' 3.64784832476320460504
Const c4 = 1.27045825245236838258           ' 1.27045825245236838258
Const c5 = 0.241780725177450611770          ' 0.241780725177450611770
Const c6 = 2.27238449892691845833e-02       ' 2.27238449892691845833e-02
Const c7 = 7.74545014278341407640e-04       ' 7.74545014278341407640e-04
Const d1 = 2.05319162663775882187           ' 2.05319162663775882187
Const d2 = 1.67638483018380384940           ' 1.67638483018380384940
Const d3 = 0.689767334985100004550          ' 0.689767334985100004550
Const d4 = 0.148103976427480074590          ' 0.148103976427480074590
Const d5 = 1.51986665636164571966e-02       ' 1.51986665636164571966e-02
Const d6 = 5.47593808499534494600e-04       ' 5.47593808499534494600e-04
Const d7 = 1.05075007164441684324e-09       ' 1.05075007164441684324e-09
'//Coefficients for P near 0 or 1.
Const e0 = 6.65790464350110377720           ' 6.65790464350110377720
Const e1 = 5.46378491116411436990           ' 5.46378491116411436990
Const e2 = 1.78482653991729133580           ' 1.78482653991729133580
Const e3 = 0.296560571828504891230          ' 0.296560571828504891230
Const e4 = 2.65321895265761230930e-02       ' 2.65321895265761230930e-02
Const e5 = 1.24266094738807843860e-03       ' 1.24266094738807843860e-03
Const e6 = 2.71155556874348757815e-05       ' 2.71155556874348757815e-05
Const e7 = 2.01033439929228813265e-07       ' 2.01033439929228813265e-07
Const f1 = 0.599832206555887937690          ' 0.599832206555887937690
Const f2 = 0.136929880922735805310          ' 0.136929880922735805310
Const f3 = 1.48753612908506148525e-02       ' 1.48753612908506148525e-02
Const f4 = 7.86869131145613259100e-04       ' 7.86869131145613259100e-04
Const f5 = 1.84631831751005468180e-05       ' 1.84631831751005468180e-05
Const f6 = 1.42151175831644588870e-07       ' 1.42151175831644588870e-07
Const f7 = 2.04426310338993978564e-15       ' 2.04426310338993978564e-15


'For poissapprox                            ' Stirling's series for Gamma(x), A001163/A001164
Const coef15 = 1# / 12#
Const coef25 = 1# / 288#
Const coef35 = -139# / 51840#
Const coef45 = -571# / 2488320#
Const coef55 = 163879# / 209018880#
Const coef65 = 5246819# / 75246796800#
Const coef75 = -534703531# / 902961561600#
Const coef1 = 2# / 3#                        ' Ramanujan's series for Gamma(x+1,x)-Gamma(x+1)/2, A065973
Const coef2 = -4# / 135#                     ' cf. http://www.whim.org/nebula/math/gammaratio.html
Const coef3 = 8# / 2835#
Const coef4 = 16# / 8505#
Const coef5 = -8992# / 12629925#
Const coef6 = -334144# / 492567075#
Const coef7 = 698752# / 1477701225#
Const coef8 = 23349012224# / 39565450299375#

Const twoThirds = 2# / 3#
Const twoFifths = 2# / 5#
Const twoSevenths = 2# / 7#
Const twoNinths = 2# / 9#
Const twoElevenths = 2# / 11#
Const twoThirteenths = 2# / 13#

'For binapprox
Const oneThird = 1# / 3#
Const twoTo27 = 134217728#                   ' 2^27

'For lngammaexpansion
Const eulers_const = 0.5772156649015328606065120901 '0.5772156649015328606065120901
Const OneMinusEulers_const = 0.4227843350984671393934879099 '0.4227843350984671393934879099
'For logfbit for small args via lngammaexpansion
Const HalfMinusEulers_const = -0.0772156649015328606065120901 '-0.0772156649015328606065120901
Const Onep25Minusln2Minuseulers_const = -0.020362845461478170023744211558177 '-0.020362845461478170023744211558177
Const FiveOver3Minusln3Minuseulers_const = -0.009161286902975885335090660355859 '-0.009161286902975885335090660355859
Const Forty7Over48Minusln4Minuseulers_const = -0.00517669268809014610764299968302 '-0.00517669268809014610764299968302
Const coeffs0Minusp3125 = 0.00996703342411321823620758332301 '0.00996703342411321823620758332301
Const coeffs0Minusp25 = 0.07246703342411321823620758332301 '0.07246703342411321823620758332301
Const coeffs0Minus1Third = -0.01086629990922011509333333333333 '-0.01086629990922011509333333333333
Const quiteSmall = 0.00000000000001

Dim hTerm As Double 'Global written to only by PBB to hold the pmf_hypergeometric value.
Dim lfbArray(29) As Double
Dim lfbArrayInitialised As Boolean
Dim coeffs(44) As Double
Dim coeffsInitialised As Boolean
Dim coeffs2(44) As Double
Dim coeffs2Initialised As Boolean

Private Sub InitAddStack(ByRef ast As TAddStack)
ast.Store = True
ast.Where = 0
ast.Stack(0).Log2Adds = 0
ast.Stack(0).value = 0#
End Sub

Private Sub DumpAddStack(ByRef ast As TAddStack)
Dim i As Integer
   Debug.Print "DumpAddStack"
   Debug.Print ast.Store
   For i = ast.Where To 0 Step -1
      Debug.Print i, ast.Stack(i).value, ast.Stack(i).Log2Adds
   Next i
   Debug.Print
End Sub

Private Sub AddValueToStack(ByRef ast As TAddStack, nextValue As Double)
If ast.Store Then
   If ast.Stack(ast.Where).Log2Adds = 0 Then
      ast.Stack(ast.Where).value = nextValue
   Else
      ast.Where = ast.Where + 1
      ast.Stack(ast.Where).value = nextValue
      ast.Stack(ast.Where).Log2Adds = 0
   End If
Else
   ast.Stack(ast.Where).value = ast.Stack(ast.Where).value + nextValue
   ast.Stack(ast.Where).Log2Adds = 1
   Do While (ast.Where > 0)
      If (ast.Stack(ast.Where).Log2Adds = ast.Stack(ast.Where - 1).Log2Adds) Then
         ast.Where = ast.Where - 1
         ast.Stack(ast.Where).value = ast.Stack(ast.Where).value + ast.Stack(ast.Where + 1).value
         ast.Stack(ast.Where).Log2Adds = ast.Stack(ast.Where).Log2Adds + 1
      Else
         Exit Do
      End If
   Loop
End If
ast.Store = Not ast.Store
End Sub

Private Function StackTotal(ByRef ast As TAddStack) As Double
Dim sum As Double, c As Double, t As Double, y As Double, i As Integer
   sum = 0#
   c = 0#
   For i = ast.Where To 0 Step -1
      y = ast.Stack(i).value - c
      t = sum + y
      c = (t - sum) - y
      sum = t
   Next i
   StackTotal = sum
End Function

Private Function TestAddValuesToStack() As Double
Dim ast As TAddStack
Dim value As Double
Dim i As Long, k As Long
Call InitAddStack(ast)
value = 1.0000000000001
For k = 1 To 10
   For i = 1 To 100000
      Call AddValueToStack(ast, value)
   Next i
Next k
'value = 10#
'Call AddValueToStack(ast, value)
'value = 20#
'Call AddValueToStack(ast, value)
'value = 30#
'Call AddValueToStack(ast, value)
'value = 40#
'Call AddValueToStack(ast, value)
'value = 50#
'Call AddValueToStack(ast, value)
'value = 60#
'Call AddValueToStack(ast, value)
'value = 70#
'Call AddValueToStack(ast, value)
'value = 80#
'Call AddValueToStack(ast, value)
'value = 90#
'Call AddValueToStack(ast, value)
'value = 100#
'Call AddValueToStack(ast, value)
TestAddValuesToStack = StackTotal(ast)
End Function

Private Sub initlfbArray()
If Not lfbArrayInitialised Then
lfbArray(0) = cf_0
lfbArray(1) = cf_1
lfbArray(2) = cf_2
lfbArray(3) = cf_3
lfbArray(4) = cf_4
lfbArray(5) = cf_5
lfbArray(6) = cf_6
lfbArray(7) = cf_7
lfbArray(8) = cf_8
lfbArray(9) = cf_9
lfbArray(10) = cf_10
lfbArray(11) = cf_11
lfbArray(12) = cf_12
lfbArray(13) = cf_13
lfbArray(14) = cf_14
lfbArray(15) = cf_15
lfbArray(16) = cf_16
lfbArray(17) = cf_17
lfbArray(18) = cf_18
lfbArray(19) = cf_19
lfbArray(20) = cf_20
lfbArray(21) = cf_21
lfbArray(22) = cf_22
lfbArray(23) = cf_23
lfbArray(24) = cf_24
lfbArray(25) = cf_25
lfbArray(26) = cf_26
lfbArray(27) = cf_27
lfbArray(28) = cf_28
lfbArray(29) = cf_29
lfbArrayInitialised = True
End If
End Sub

Private Sub initCoeffs()
If Not coeffsInitialised Then
'// for i < UBound coeffs, coeffs[i] holds (zeta(i+2)-1)/(i+2), coeffs[UBound coeffs] holds (zeta(UBound coeffs+2)-1)
coeffs(0)=0.32246703342411321824    '0.32246703342411321824
coeffs(1)=6.7352301053198095133e-02 '6.7352301053198095133e-02
coeffs(2)=2.0580808427784547879e-02 '2.0580808427784547879e-02
coeffs(3)=7.3855510286739852663e-03 '7.3855510286739852663e-03
coeffs(4)=2.8905103307415232858e-03 '2.8905103307415232858e-03
coeffs(5)=1.1927539117032609771e-03 '1.1927539117032609771e-03
coeffs(6)=5.0966952474304242234e-04 '5.0966952474304242234e-04
coeffs(7)=2.2315475845357937976e-04 '2.2315475845357937976e-04
coeffs(8)=9.9457512781808533715e-05 '9.9457512781808533715e-05
coeffs(9)=4.4926236738133141700e-05 '4.4926236738133141700e-05
coeffs(10)=2.0507212775670691553e-05 '2.0507212775670691553e-05
coeffs(11)=9.4394882752683959040e-06 '9.4394882752683959040e-06
coeffs(12)=4.3748667899074878042e-06 '4.3748667899074878042e-06
coeffs(13)=2.0392157538013662368e-06 '2.0392157538013662368e-06
coeffs(14)=9.5514121304074198329e-07 '9.5514121304074198329e-07
coeffs(15)=4.4924691987645660433e-07 '4.4924691987645660433e-07
coeffs(16)=2.1207184805554665869e-07 '2.1207184805554665869e-07
coeffs(17)=1.0043224823968099609e-07 '1.0043224823968099609e-07
coeffs(18)=4.7698101693639805658e-08 '4.7698101693639805658e-08
coeffs(19)=2.2711094608943164910e-08 '2.2711094608943164910e-08
coeffs(20)=1.0838659214896954091e-08 '1.0838659214896954091e-08
coeffs(21)=5.1834750419700466551e-09 '5.1834750419700466551e-09
coeffs(22)=2.4836745438024783172e-09 '2.4836745438024783172e-09
coeffs(23)=1.1921401405860912074e-09 '1.1921401405860912074e-09
coeffs(24)=5.7313672416788620133e-10 '5.7313672416788620133e-10
coeffs(25)=2.7595228851242331452e-10 '2.7595228851242331452e-10
coeffs(26)=1.3304764374244489481e-10 '1.3304764374244489481e-10
coeffs(27)=6.4229645638381000221e-11 '6.4229645638381000221e-11
coeffs(28)=3.1044247747322272762e-11 '3.1044247747322272762e-11
coeffs(29)=1.5021384080754142171e-11 '1.5021384080754142171e-11
coeffs(30)=7.2759744802390796625e-12 '7.2759744802390796625e-12
coeffs(31)=3.5277424765759150836e-12 '3.5277424765759150836e-12
coeffs(32)=1.7119917905596179086e-12 '1.7119917905596179086e-12
coeffs(33)=8.3153858414202848198e-13 '8.3153858414202848198e-13
coeffs(34)=4.0422005252894400655e-13 '4.0422005252894400655e-13
coeffs(35)=1.9664756310966164904e-13 '1.9664756310966164904e-13
coeffs(36)=9.5736303878385557638e-14 '9.5736303878385557638e-14
coeffs(37)=4.6640760264283742246e-14 '4.6640760264283742246e-14
coeffs(38)=2.2737369600659723206e-14 '2.2737369600659723206e-14
coeffs(39)=1.1091399470834522017e-14 '1.1091399470834522017e-14
coeffs(40)=5.4136591567253631315e-15 '5.4136591567253631315e-15
coeffs(41)=2.6438800178609949985e-15 '2.6438800178609949985e-15
coeffs(42)=1.2918959062789967293811764562316e-15 '1.2918959062789967293811764562316e-15
coeffs(43)=6.3159355041984485676779394847024e-16 '6.3159355041984485676779394847024e-16
coeffs(44)=1.421085482803160676983430580383e-14 '1.421085482803160676983430580383e-14
coeffsInitialised = True
End If
End Sub
Private Sub initCoeffs2()
If Not coeffs2Initialised Then
'// coeffs[i] holds (zeta(i+2)-1)/(i+2) - (i+5)/(i+2)/(i+1)*2^(-i-3)
coeffs2(0)=9.96703342411321823620758332301e-3 '9.96703342411321823620758332301e-3
coeffs2(1)=4.85230105319809513323333333333e-3 '4.85230105319809513323333333333e-3
coeffs2(2)=2.35164176111788121233425746862e-3 '2.35164176111788121233425746862e-3
coeffs2(3)=1.1355510286739852662730972914e-3 '1.1355510286739852662730972914e-3
coeffs2(4)=5.4676033074152328575298829848e-4 '5.4676033074152328575298829848e-4
coeffs2(5)=2.6269438789373716759012616902e-4 '2.6269438789373716759012616902e-4
coeffs2(6)=1.2601997117161385090708338501e-4 '1.2601997117161385090708338501e-4
coeffs2(7)=6.03943417869127130948e-5 '6.03943417869127130948e-5
coeffs2(8)=2.89279988929196448257e-5 '2.89279988929196448257e-5
coeffs2(9)=1.38537935563149598820e-5 '1.38537935563149598820e-5
coeffs2(10)=6.63558635521614609862e-6 '6.63558635521614609862e-6
coeffs2(11)=3.17947224962737026300e-6 '3.17947224962737026300e-6
coeffs2(12)=1.52432377823166362836e-6 '1.52432377823166362836e-6
coeffs2(13)=7.31319548444223379639e-7 '7.31319548444223379639e-7
coeffs2(14)=3.51147479316783649952e-7 '3.51147479316783649952e-7
coeffs2(15)=1.68754473874618369035e-7 '1.68754473874618369035e-7
coeffs2(16)=8.11753732546888155551e-8 '8.11753732546888155551e-8
coeffs2(17)=3.90847775936649142159e-8 '3.90847775936649142159e-8
coeffs2(18)=1.88369052760822398681e-8 '1.88369052760822398681e-8
coeffs2(19)=9.08717580313959348175e-9 '9.08717580313959348175e-9
coeffs2(20)=4.38794008336117216467e-9 '4.38794008336117216467e-9
coeffs2(21)=2.12078578473653627963e-9 '2.12078578473653627963e-9
coeffs2(22)=1.02595225309999020577e-9 '1.02595225309999020577e-9
coeffs2(23)=4.96752618206533915776e-10 '4.96752618206533915776e-10
coeffs2(24)=2.40726205228207715756e-10 '2.40726205228207715756e-10
coeffs2(25)=1.16751848407213311847e-10 '1.16751848407213311847e-10
coeffs2(26)=5.66693373586358102002e-11 '5.66693373586358102002e-11
coeffs2(27)=2.75272781658498271913e-11 '2.75272781658498271913e-11
coeffs2(28)=1.33812334011666457419e-11 '1.33812334011666457419e-11
coeffs2(29)=6.50929603319331702812e-12 '6.50929603319331702812e-12
coeffs2(30)=3.16857905287746826547e-12 '3.16857905287746826547e-12
coeffs2(31)=1.54339039998043529180e-12 '1.54339039998043529180e-12
coeffs2(32)=7.52239805801019839357e-13 '7.52239805801019839357e-13
coeffs2(33)=3.66855576849641617566e-13 '3.66855576849641617566e-13
coeffs2(34)=1.79011840661361776213e-13 '1.79011840661361776213e-13
coeffs2(35)=8.73989502840846835258e-14 '8.73989502840846835258e-14
coeffs2(36)=4.26932273880725309600e-14 '4.26932273880725309600e-14
coeffs2(37)=2.08656067727432658676e-14 '2.08656067727432658676e-14
coeffs2(38)=1.02026669800712891581e-14 '1.02026669800712891581e-14
coeffs2(39)=4.99113012967463749398e-15 '4.99113012967463749398e-15
coeffs2(40)=2.44274876330334144841e-15 '2.44274876330334144841e-15
coeffs2(41)=1.19604099925791156327e-15 '1.19604099925791156327e-15
coeffs2(42)=5.85859784064943690566e-16 '5.85859784064943690566e-16
coeffs2(43)=2.87087981566462948467e-16 '2.87087981566462948467e-16
coeffs2(44)=1.40735520163755175636e-16 '1.40735520163755175636e-16
'coeffs2(45)=6.90166402588687226293e-17 '6.90166402588687226293e-17
'coeffs2(46)=3.38578655511664755968e-17 '3.38578655511664755968e-17
'coeffs2(47)=1.66155827667479862403e-17 '1.66155827667479862403e-17
'coeffs2(48)=8.15674061694190990538e-18 '8.15674061694190990539e-18
'coeffs2(49)=4.00551052932053968016e-18 '4.00551052932053968016e-18
'coeffs2(50)=1.96758982791540465538e-18 '1.96758982791540465538e-18
'coeffs2(51)=9.66812504275437439100e-19 '9.66812504275437439100e-19
'coeffs2(52)=4.75200281648237941674e-19 '4.75200281648237941674e-19
'coeffs2(53)=2.33632791481571906191e-19 '2.33632791481571906191e-19
'coeffs2(54)=1.14897269222195236258e-19 '1.14897269222195236258e-19
'coeffs2(55)=5.65198125116716026715e-20 '5.651981251167160267215-20
'coeffs2(56)=2.78101464727381599671e-20 '2.78101464727381599671e-20
'coeffs2(57)=1.36871811383630687470e-20 '1.36871811383630687470e-20
'coeffs2(58)=6.73797960341042295487e-21 '6.73797960341042295487e-21
'coeffs2(59)=3.31777713997525851631e-21 '3.31777713997525851631e-21
'coeffs2(60)=1.63404346465554253504e-21 '1.63404346465554253504e-21
'coeffs2(61)=8.04963210512578619264e-22 '8.04963210512578619264e-22
'coeffs2(62)=3.96626538798230823888e-22 '3.96626538798230823888e-22
'coeffs2(63)=1.95469141675562808544e-22 '1.95469141675562808544e-22
'coeffs2(64)=9.63524657953850607531e-23 '9.63524657953850607531e-23
'coeffs2(65)=4.75043353504700102959e-23 '4.75043353504700102959e-23
'coeffs2(66)=2.34254063551981428148e-23 '2.34254063551981428148e-23
'coeffs2(67)=1.15537315908688659068e-23 '1.15537315908688659068e-23
'coeffs2(68)=5.69949705709986050247e-24 '5.69949705709986050247e-24
'coeffs2(69)=2.81208121322037769865e-24 '2.81208121322037769865e-24
'coeffs2(70)=1.38769580071555495253e-24 '1.38769580071555495253e-24
coeffs2Initialised = True
End If
End Sub

Private Function ec() As Double
ec = eulers_const
End Function

Private Function logcfdersum(ByVal x As Double, ByVal i As Double, ByVal d As Double, Optional ByVal derivs As Integer = 0) As Double
'// Calculation of logcfdersum(x,i,d,derivs) via summation, where derivs is the number of derivatives wrt x required of the normal logcf(x,i,d) function.
'// Will become hopelessly slow for x >= 0.5
Dim tot As Double, addon As Double, y As Double, nd As Double, k As Integer
tot = 1#
y = x / (x - 1#)
nd = (derivs + 1#) * d
addon = y * nd / (i + nd)
While Abs(addon) > Abs(0.000000000000001 * tot)
    tot = tot + addon
    nd = nd + d
    addon = addon * y * nd / (i + nd)
Wend
logcfdersum = (tot + addon) / ((i + derivs * d) * (1# - x) ^ (derivs + 1))
For k = 2 To derivs
   logcfdersum = logcfdersum * k
Next k
End Function

Private Function deriv2cf(ByVal x As Double, ByVal i As Double, ByVal d As Double) As Double
'// Accurate calculation of derivative of logcf(x,i,d) wrt x, when x is small in absolute value.
Dim n As Double, j As Double, tot As Double, xtojm1 As Double, addon As Double
n = i + 2# * d
tot = 2# / n
j = 2#
n = n + d
xtojm1 = x
addon = j * (j + 1#) * xtojm1 / n
While Abs(addon) > Abs(0.000000000000001 * tot)
    tot = tot + addon
    xtojm1 = x * xtojm1
    j = j + 1#
    n = n + d
    addon = j * (j + 1#) * xtojm1 / n
Wend
deriv2cf = tot + addon
End Function


Private Function derivcf(ByVal x As Double, ByVal i As Double, ByVal d As Double) As Double
'// Accurate calculation of derivative of logcf(x,i,d) wrt x, when x is small in absolute value.
Dim n As Double, j As Double, tot As Double, xtojm1 As Double, addon As Double
n = i + d
tot = 1# / n
j = 2#
n = n + d
xtojm1 = x
addon = j * xtojm1 / n
While Abs(addon) > Abs(1E-16 * tot)
    tot = tot + addon
    xtojm1 = x * xtojm1
    j = j + 1#
    n = n + d
    addon = j * xtojm1 / n
Wend
derivcf = tot + addon
End Function

Private Function Min(ByVal x As Double, ByVal y As Double) As Double
   If x < y Then
      Min = x
   Else
      Min = y
   End If
End Function
Private Function Max(ByVal x As Double, ByVal y As Double) As Double
   If x > y Then
      Max = x
   Else
      Max = y
   End If
End Function

Private Function expm1old(ByVal x As Double) As Double
'// Accurate calculation of exp(x)-1, particularly for small x.
'// Uses a variation of the standard continued fraction for tanh(x) see A&S 4.5.70.
  If (Abs(x) < 2) Then
     Dim a1 As Double, a2 As Double, b1 As Double, b2 As Double, c1 As Double, x2 As Double
     a1 = 24#
     b1 = 2# * (12# - x * (6# - x))
     x2 = x * x * 0.25
     a2 = 8# * (15# + x2)
     b2 = 120# - x * (60# - x * (12# - x))
     c1 = 7#

     While ((Abs(a2 * b1 - a1 * b2) > Abs(cfSmall * b1 * a2)))

       a1 = c1 * a2 + x2 * a1
       b1 = c1 * b2 + x2 * b1
       c1 = c1 + 2#

       a2 = c1 * a1 + x2 * a2
       b2 = c1 * b1 + x2 * b2
       c1 = c1 + 2#
       If (b2 > scalefactor) Then
         a1 = a1 * scalefactor2
         b1 = b1 * scalefactor2
         a2 = a2 * scalefactor2
         b2 = b2 * scalefactor2
       End If
     Wend

     expm1old = x * a2 / b2
  Else
     expm1old = Exp(x) - 1#
  End If

End Function

Private Function expm1(ByVal x As Double) As Double
'// Accurate calculation of exp(x)-1, particularly for small x.
'// Based on NR approach to solving log(1+result) = x
Dim y0 As Double, a2 As Double, b1 As Double, b2 As Double, c1 As Double, x2 As Double
  y0 = Exp(x) - 1#
  If Abs(x) < 2 Then
     If y0 = 0# Then
        expm1 = x
     Else
        expm1 = y0 - (Log(1# + y0) - x) * (1# + y0)
     End If
  Else
     expm1 = y0
  End If
End Function

Private Function logcf(ByVal x As Double, ByVal i As Double, ByVal d As Double) As Double
'// Continued fraction for calculation of 1/i + x/(i+d) + x*x/(i+2*d) + x*x*x/(i+3d) + ...
Dim a1 As Double, a2 As Double, b1 As Double, b2 As Double, c1 As Double, c2 As Double, c3 As Double, c4 As Double
     c1 = 2# * d
     c2 = i + d
     c4 = c2 + d
     a1 = c2
     b1 = i * (c2 - i * x)
     b2 = d * d * x
     a2 = c4 * c2 - b2
     b2 = c4 * b1 - i * b2

     While ((Abs(a2 * b1 - a1 * b2) > Abs(cfVSmall * b1 * a2)))

       c3 = c2 * c2 * x
       c2 = c2 + d
       c4 = c4 + d
       a1 = c4 * a2 - c3 * a1
       b1 = c4 * b2 - c3 * b1

       c3 = c1 * c1 * x
       c1 = c1 + d
       c4 = c4 + d
       a2 = c4 * a1 - c3 * a2
       b2 = c4 * b1 - c3 * b2
       If (b2 > scalefactor) Then
         a1 = a1 * scalefactor2
         b1 = b1 * scalefactor2
         a2 = a2 * scalefactor2
         b2 = b2 * scalefactor2
       ElseIf (b2 < scalefactor2) Then
         a1 = a1 * scalefactor
         b1 = b1 * scalefactor
         a2 = a2 * scalefactor
         b2 = b2 * scalefactor
       End If
     Wend
     logcf = a2 / b2
End Function

Private Function logcfplusderiv(ByVal x As Double, ByVal i As Double, ByVal d As Double) As Double
'// Continued fraction type calculation of derivative of 1/i + x/(i+d) + x*x/(i+2*d) + x*x*x/(i+3d) + ...
Dim a1 As Double, a2 As Double, b1 As Double, b2 As Double, a1dash As Double, a2dash As Double, b1dash As Double, b2dash As Double, c1 As Double, c2 As Double, c3 As Double, c4 As Double, c5 As Double
     c1 = 2# * d
     c2 = i + d
     c4 = c2 + d
     a1 = c2
     a1dash = 0#
     b1 = i * (c2 - i * x)
     b1dash = -i * i
     b2 = d * d * x
     a2 = c4 * c2 - b2
     a2dash = -d * d
     b2 = c4 * b1 - i * b2
     b2dash = c4 * b1dash + i * a2dash

     While ((Abs(a2 * b1 - a1 * b2) > Abs(cfVSmall * b1 * a2)))

       c5 = c2 * c2
       c3 = c5 * x
       c2 = c2 + d
       c4 = c4 + d
       a1dash = c4 * a2dash - c3 * a1dash - c5 * a1
       b1dash = c4 * b2dash - c3 * b1dash - c5 * b1
       a1 = c4 * a2 - c3 * a1
       b1 = c4 * b2 - c3 * b1

       c5 = c1 * c1
       c3 = c5 * x
       c1 = c1 + d
       c4 = c4 + d
       a2dash = c4 * a1dash - c3 * a2dash - c5 * a2
       b2dash = c4 * b1dash - c3 * b2dash - c5 * b2
       a2 = c4 * a1 - c3 * a2
       b2 = c4 * b1 - c3 * b2
       If (b2 > scalefactor) Then
         a1 = a1 * scalefactor2
         b1 = b1 * scalefactor2
         a2 = a2 * scalefactor2
         b2 = b2 * scalefactor2
         a1dash = a1dash * scalefactor2
         b1dash = b1dash * scalefactor2
         a2dash = a2dash * scalefactor2
         b2dash = b2dash * scalefactor2
       ElseIf (b2 < scalefactor2) Then
         a1 = a1 * scalefactor
         b1 = b1 * scalefactor
         a2 = a2 * scalefactor
         b2 = b2 * scalefactor
         a1dash = a1dash * scalefactor
         b1dash = b1dash * scalefactor
         a2dash = a2dash * scalefactor
         b2dash = b2dash * scalefactor
       End If
     Wend
     c5 = c2 * c2
     c3 = c5 * x
     c4 = c4 + d
     a1dash = c4 * a2dash - c3 * a1dash - c5 * a1
     b1dash = c4 * b2dash - c3 * b1dash - c5 * b1
     a1 = c4 * a2 - c3 * a1
     b1 = c4 * b2 - c3 * b1
     logcfplusderiv = (a1dash * b1 - a1 * b1dash) / b1 ^ 2
End Function

Private Function log0Old(ByVal x As Double) As Double
'//Accurate calculation of log(1+x), particularly for small x.
   Dim term As Double
   If (Abs(x) > 0.5) Then
      log0Old = Log(1# + x)
   Else
     term = x / (2# + x)
     log0Old = 2# * term * logcf(term * term, 1#, 2#)
   End If
End Function

Private Function log0(ByVal x As Double) As Double
'//Accurate and quicker calculation of log(1+x), particularly for small x. Code from Wolfgang Ehrhardt.
   Dim y As Double
   If x > 4# Then
      log0 = Log(1# + x)
   Else
      y = 1# + x
      If y = 1# Then
         log0 = x
      Else
         log0 = Log(y) + (x - (y - 1#)) / y
      End If
   End If
End Function

Private Function lcc(ByVal x As Double) As Double
'//Accurate calculation of log(1+x)-x, particularly for small x.
   Dim term As Double, y  As Double
   If (Abs(x) < 0.01) Then
      term = x / (2# + x)
      y = term * term
      lcc = term * ((((2# / 9# * y + 2# / 7#) * y + 0.4) * y + 2# / 3#) * y - x)
   ElseIf (x < minLog1Value Or x > 1#) Then
      lcc = Log(1# + x) - x
   Else
      term = x / (2# + x)
      y = term * term
      lcc = term * (2# * y * logcf(y, 3#, 2#) - x)
   End If
End Function

Private Function log1(ByVal x As Double) As Double
'//Accurate calculation of log(1+x)-x, particularly for small x.
   Dim term As Double, y  As Double
   If (Abs(x) < 0.01) Then
      term = x / (2# + x)
      y = term * term
      log1 = term * ((((2# / 9# * y + 2# / 7#) * y + 0.4) * y + 2# / 3#) * y - x)
   ElseIf (x < minLog1Value Or x > 1#) Then
      log1 = Log(1# + x) - x
   Else
      term = x / (2# + x)
      y = term * term
      log1 = term * (2# * y * logcf(y, 3#, 2#) - x)
   End If
End Function

Private Function logfbitdif(ByVal x As Double) As Double
'//Calculation of logfbit(x)-logfbit(1+x). x must be > -1.
  Dim y As Double, y2 As Double
  If x < -0.65 Then
     logfbitdif = (x + 1.5) * log0(1# / (x + 1#)) - 1#
  Else
     y2 = (2# * x + 3#) ^ -2
     logfbitdif = y2 * logcf(y2, 3#, 2#)
  End If
End Function

Private Function logfbita(ByVal x As Double) As Double
'//Error part of Stirling's formula where log(x!) = log(sqrt(twopi))+(x+0.5)*log(x+1)-(x+1)+logfbita(x).
'//Are we ever concerned about the relative error involved in this function? I don't think so.
  Dim x1 As Double, x2 As Double, x3 As Double
  If (x >= 100000000#) Then
     logfbita = lfbc1 / (x + 1#)
  ElseIf (x >= 6#) Then                      ' Abramowitz & Stegun's series 6.1.41
     x1 = x + 1#
     x2 = 1# / (x1 * x1)
     x3 = x2 * (lfbc6 - x2 * (lfbc7 - x2 * (lfbc8 - x2 * lfbc9)))
     x3 = x2 * (lfbc4 - x2 * (lfbc5 - x3))
     x3 = x2 * (lfbc2 - x2 * (lfbc3 - x3))
     logfbita = lfbc1 * (1# - x3) / x1
  ElseIf (x = 0#) Then
     logfbita = lfb_0
  ElseIf (x = 1#) Then
     logfbita = lfb_1
  ElseIf (x = 2#) Then
     logfbita = lfb_2
  ElseIf (x = 3#) Then
     logfbita = lfb_3
  ElseIf (x = 4#) Then
     logfbita = lfb_4
  ElseIf (x = 5#) Then
     logfbita = lfb_5
  ElseIf (x > -1#) Then
     x1 = x
     x2 = 0#
     While (x1 < 6#)
        x2 = x2 + logfbitdif(x1)
        x1 = x1 + 1#
     Wend
     logfbita = x2 + logfbita(x1)
  Else
     logfbita = 1E+308
  End If
End Function

Private Function logfbitb(ByVal x As Double) As Double
    Dim lgam As Double
    Dim i As Integer
    Dim m As Double
    Dim big As Boolean
    Call initCoeffs
    If x <= 0.5 Then
       m = 0#
    ElseIf x <= 1.5 Then
       m = 1#
    ElseIf x <= 2.5 Then
       m = 2#
    Else
       m = 3#
    End If
    x = x - m
    i = UBound(coeffs)
    lgam = coeffs(i) * logcf(-x / 2#, i + 2#, 1#)
    For i = UBound(coeffs) - 1 To 2 Step -1
       lgam = (coeffs(i) - x * lgam)
    Next i
    If m = 3# Then
       lgam = (coeffs(1) - x * lgam)
       logfbitb = ((x * x * (coeffs0Minusp25 - x * lgam) - (x + 3.5) * log1(x / 4#) + log1(x / 2#) + log1(x / 3#)) + Forty7Over48Minusln4Minuseulers_const * x) + lfb_3
    ElseIf m = 2# Then
       lgam = (coeffs(1) - x * lgam)
       logfbitb = ((x * x * (coeffs0Minus1Third - x * lgam) - (x + 2.5) * log1(x / 3#) + log1(x / 2#)) + FiveOver3Minusln3Minuseulers_const * x) + lfb_2
    ElseIf m = 1# Then
       'lgam = (coeffs(1) - x * lgam)
       'logfbitb = ((x * x * (coeffs(0) - 0.5 - x * lgam) - (x + 1.5) * log1(x / 2#)) + Onep25Minusln2Minuseulers_const * x) + lfb_1
       'logfbit = ((coeffs0Minusp3125 - (lgam - 0.125 * (1# - (x + 1.5) * logcf(-x / 2#, 3#, 1#))) * x) * x + Onep25Minusln2Minuseulers_const) * x + lfb_1
       'logfbitb = ((coeffs0Minusp3125 - (coeffs(1) - 0.125 - x * lgam + 0.125 * (x + 1.5) * logcf(-x / 2#, 3#, 1#)) * x) * x + Onep25Minusln2Minuseulers_const) * x + lfb_1
       logfbitb = ((coeffs0Minusp3125 - (coeffs(1) - 0.0625 - x * (lgam - 1# / 24# + 0.0625 * (x + 1.5) * logcf(-x / 2#, 4#, 1#))) * x) * x + Onep25Minusln2Minuseulers_const) * x + lfb_1
    Else
       'lgam = (coeffs(1) - x * lgam)
       'logfbitb = ((x * x * (coeffs(0) - 1# - x * lgam) - (x + 1.5) * log1(x)) + HalfMinusEulers_const * x) + lfb_0
       'logfbitb = ((coeffs0Minusp25 - ((x + 1.5) * logcf(-x, 3#, 1#) - 0.5 + lgam) * x) * x + HalfMinusEulers_const) * x + lfb_0
       'logfbitb = ((coeffs0Minusp25 - (x * (1# / 3# - (x + 1.5) * logcf(-x, 4#, 1#)) + lgam) * x) * x + HalfMinusEulers_const) * x + lfb_0
       logfbitb = ((coeffs0Minusp25 - (coeffs(1) + (x * (x + 1.5) * logcf(-x, 5#, 1#) - (6# * x + 1#) / 24# - lgam) * x) * x) * x + HalfMinusEulers_const) * x + lfb_0
    End If
End Function

Private Function logfbit(ByVal x As Double) As Double
'//Calculates log of x factorial - log(sqrt(2*pi)) +(x+1) -(x+0.5)*log(x+1)
'//using the error part of Stirling's formula (see Abramowitz & Stegun's series 6.1.41)
'//and Stieltjes' continued fraction for the gamma function.
'//For x < 1.5, uses expansion of log(x!) and log((x+1)!) from Abramowitz & Stegun's series 6.1.33
'//We are primarily concerned about the absolute error in this function.
'//Due to cancellation errors in calculating 1+x as x tends to -1, the function loses accuracy and should not be used!
  Dim x1 As Double, x2 As Double, x3 As Double
  If (x >= 6#) Then
     x1 = x + 1#
     If (x >= 1000#) Then
        If (x >= 100000000#) Then
           x3 = 0#
        Else
           x2 = 1# / (x1 * x1)
           x3 = x2 * (lfbc2 - x2 * lfbc3)
        End If
     Else
        x2 = 1# / (x1 * x1)
        If x >= 40# Then
           x3 = 0#
        ElseIf x >= 15# Then
           x3 = x2 * (lfbc6 - x2 * lfbc7)
        Else
           x3 = x2 * (lfbc6 - x2 * (lfbc7 - x2 * (lfbc8 - x2 * lfbc9)))
        End If
        x3 = x2 * (lfbc4 - x2 * (lfbc5 - x3))
        x3 = x2 * (lfbc2 - x2 * (lfbc3 - x3))
     End If
     logfbit = lfbc1 * (1# - x3) / x1
     'logfbit = (1# - x3) / (12# * x1)
  ElseIf (x = 0#) Then
     logfbit = lfb_0
  ElseIf (x = 1#) Then
     logfbit = lfb_1
  ElseIf (x = 2#) Then
     logfbit = lfb_2
  ElseIf (x = 3#) Then
     logfbit = lfb_3
  ElseIf (x = 4#) Then
     logfbit = lfb_4
  ElseIf (x = 5#) Then
     logfbit = lfb_5
  ElseIf x > 1.5 Then
     x1 = x + 1#
     If x >= 2.5 Then
        'x2 = 0.25 * ((Sqr(x1 * x1 + 81#) - x1) + 81# / (x1 + Sqr(x1 * x1 + 90.25)))
        x2 = 40.5 / (x1 + Sqr(x1 * x1 + 81#))
     Else
        'x2 = 0.25 * ((Sqr(x1 * x1 + 225#) - x1) + 225# / (x1 + Sqr(x1 * x1 + 240.25)))
        x2 = 112.5 / (x1 + Sqr(x1 * x1 + 225#))
        x2 = cf_27 / (x1 + cf_28 / (x1 + cf_29 / (x1 + x2)))
        x2 = cf_24 / (x1 + cf_25 / (x1 + cf_26 / (x1 + x2)))
        x2 = cf_21 / (x1 + cf_22 / (x1 + cf_23 / (x1 + x2)))
        x2 = cf_18 / (x1 + cf_19 / (x1 + cf_20 / (x1 + x2)))
     End If
     x2 = cf_15 / (x1 + cf_16 / (x1 + cf_17 / (x1 + x2)))
     x2 = cf_12 / (x1 + cf_13 / (x1 + cf_14 / (x1 + x2)))
     x2 = cf_9 / (x1 + cf_10 / (x1 + cf_11 / (x1 + x2)))
     x2 = cf_6 / (x1 + cf_7 / (x1 + cf_8 / (x1 + x2)))
     x2 = cf_3 / (x1 + cf_4 / (x1 + cf_5 / (x1 + x2)))
     'logfbit = cf_0 / (x1 + cf_1 / (x1 + cf_2 / (x1 + x2)))
     logfbit = 1# / (12# * (x1 + cf_1 / (x1 + cf_2 / (x1 + x2))))
  'ElseIf (x = 1.5) Then
  '   logfbit = 3.316287351993628748511050974106e-02   ' 3.316287351993628748511050974106e-02
  ElseIf (x = 0.5) Then
     logfbit = logfbit0p5                             ' 5.481412105191765389613870234839e-02
  ElseIf (x = -0.5) Then
     logfbit = 0.15342640972002734529138393927091     ' 0.15342640972002734529138393927091
  ElseIf x >= -0.65 Then
    Dim lgam As Double
    Dim i As Integer
    If x <= 0# Then
       Call initCoeffs
       i = UBound(coeffs)
       lgam = coeffs(i) * logcf(-x / 2#, i + 2#, 1#)
       For i = UBound(coeffs) - 1 To 1 Step -1
          lgam = (coeffs(i) - x * lgam)
       Next i
       logfbit = ((coeffs0Minusp25 - (x * (1# / 3# - (x + 1.5) * logcf(-x, 4#, 1#)) + lgam) * x) * x + HalfMinusEulers_const) * x + lfb_0
    ElseIf x <= 1.56 Then
       x = x - 1#
       Call initCoeffs2
       i = UBound(coeffs2) + 3
       lgam = ((x + 2.5) * logcf(-x / 2#, i, 1#) - (2# / (i - 1#))) * (2# ^ -i) + (3# ^ -i) * logcf(-x / 3#, i, 1#)
       For i = UBound(coeffs2) To 0 Step -1
          lgam = (coeffs2(i) - x * lgam)
       Next i
       logfbit = (x * lgam + Onep25Minusln2Minuseulers_const) * x + lfb_1
    ElseIf x <= 2.5 Then
       x = x - 2#
       Call initCoeffs
       i = UBound(coeffs)
       lgam = coeffs(i) * logcf(-x / 2#, i + 2#, 1#)
       For i = UBound(coeffs) - 1 To 1 Step -1
          lgam = (coeffs(i) - x * lgam)
       Next i
       logfbit = ((x * x * (coeffs0Minus1Third - x * lgam) - (x + 2.5) * log1(x / 3#) + log1(x / 2#)) + FiveOver3Minusln3Minuseulers_const * x) + lfb_2
    Else
       x = x - 3#
       Call initCoeffs
       i = UBound(coeffs)
       lgam = coeffs(i) * logcf(-x / 2#, i + 2#, 1#)
       For i = UBound(coeffs) - 1 To 1 Step -1
          lgam = (coeffs(i) - x * lgam)
       Next i
       logfbit = ((x * x * (coeffs0Minusp25 - x * lgam) - (x + 3.5) * log1(x / 4#) + log1(x / 2#) + log1(x / 3#)) + Forty7Over48Minusln4Minuseulers_const * x) + lfb_3
    End If
  ElseIf x > -1# Then
    logfbit = logfbitdif(x) + logfbit(x + 1#)
  Else
     logfbit = 1E+308
  End If
End Function

Private Function lfbaccdif1(ByVal a As Double, ByVal b As Double) As Double
'//Calculates logfbit(b)-logfbit(a+b) accurately for a > 0 & b >= 0. Reasonably accurate for a >=0 & b < 0.
Dim x1 As Double, x2 As Double, x3 As Double, y1 As Double, y2 As Double, y3 As Double
Dim acc As Double, i As Integer, Start As Integer, s1 As Double, s2 As Double, tx As Double, ty As Double
  If a < 0# Then
     lfbaccdif1 = -lfbaccdif1(-a, b + a)
  ElseIf (b >= 8#) Then
     y1 = b + 1#
     y2 = y1 ^ -2
     x1 = a + b + 1#
     x2 = x1 ^ -2
     x3 = x2 * lfbc9
     y3 = y2 * lfbc9
     acc = x2 * (a * (x1 + y1) * y3)
     x3 = x2 * (lfbc8 - x3)
     y3 = y2 * (lfbc8 - y3)
     acc = x2 * (a * (x1 + y1) * y3 - acc)
     x3 = x2 * (lfbc7 - x3)
     y3 = y2 * (lfbc7 - y3)
     acc = x2 * (a * (x1 + y1) * y3 - acc)
     x3 = x2 * (lfbc6 - x3)
     y3 = y2 * (lfbc6 - y3)
     acc = x2 * (a * (x1 + y1) * y3 - acc)
     x3 = x2 * (lfbc5 - x3)
     y3 = y2 * (lfbc5 - y3)
     acc = x2 * (a * (x1 + y1) * y3 - acc)
     x3 = x2 * (lfbc4 - x3)
     y3 = y2 * (lfbc4 - y3)
     acc = x2 * (a * (x1 + y1) * y3 - acc)
     x3 = x2 * (lfbc3 - x3)
     y3 = y2 * (lfbc3 - y3)
     acc = x2 * (a * (x1 + y1) * y3 - acc)
     x3 = x2 * (lfbc2 - x3)
     y3 = y2 * (lfbc2 - y3)
     acc = x2 * (a * (x1 + y1) * y3 - acc)
     'lfbaccdif1 = lfbc1 * (a * (1# - y3) - y1 * acc) / (x1 * y1)
     lfbaccdif1 = (a * (1# - y3) - y1 * acc) / (12# * x1 * y1)
  ElseIf b >= 1.7 Then
     y1 = b + 1#
     x1 = a + b + 1#
     If b >= 3# Then
        Start = 17
     Else
        Start = 29
     End If
     s1 = (0.5 * (Start + 1#)) ^ 2
     s2 = (0.5 * (Start + 1.5)) ^ 2
     ty = y1 * Sqr(1# + s1 * (y1 ^ -2))
     tx = x1 * Sqr(1# + s1 * (x1 ^ -2))
     y2 = ty - y1
     x2 = tx - x1
     acc = a * (1# - (2# * y1 + a) / (tx + ty))
     'Seems to work better without the next 2 lines. - Not with modification to s2
     ty = y1 * Sqr(1# + s2 * (y1 ^ -2))
     tx = x1 * Sqr(1# + s2 * (x1 ^ -2))
     acc = 0.25 * (acc + s1 / ((y1 + ty) * (x1 + tx)) * a * (1# + (2# * y1 + a) / (tx + ty)))
     y2 = 0.25 * (y2 + s1 / (y1 + ty))
     x2 = 0.25 * (x2 + s1 / (x1 + tx))
     Call initlfbArray
     For i = Start To 1 Step -1
        acc = lfbArray(i) * (a - acc) / ((x1 + x2) * (y1 + y2))
        y2 = lfbArray(i) / (y1 + y2)
        x2 = lfbArray(i) / (x1 + x2)
     Next i
     lfbaccdif1 = cf_0 * (a - acc) / ((x1 + x2) * (y1 + y2))
     'lfbaccdif1 = (a - acc) / (12# * (x1 + x2) * (y1 + y2))
  ElseIf b > -1# Then
    Dim scale2 As Double, scale3 As Double
    If b < -0.66 Then
       If a > 1# Then
          lfbaccdif1 = logfbitdif(b) + lfbaccdif1(a - 1#, b + 1#)
          Exit Function
       ElseIf a = 1# Then
          lfbaccdif1 = logfbitdif(b)
          Exit Function
       Else
          s2 = a * log0(1# / (b + 1# + a))
          s1 = logfbitdif(b + a)
          If s1 > s2 Then
             s1 = (b + 1.5) * log0(a / ((b + 1#) * (b + 2# + a))) - s2
          Else
             s2 = s1
             s1 = (logfbitdif(b) - s1)
          End If
          If s1 > 0.1 * s2 Then
             lfbaccdif1 = s1 + lfbaccdif1(a, b + 1#)
             Exit Function
          End If
       End If
    End If
    Call initCoeffs2
    If b + a > 2 Then
       s1 = lfbaccdif1(b + a - 1.75, 1.75)
       a = 1.75 - b
    Else
       s1 = 0#
    End If
    y1 = b - 1#
    x1 = y1 + a
    i = UBound(coeffs2) + 3
    scale2 = 2# ^ -i
    scale3 = 3# ^ -i
    'y2 = ((y1 + 2.5) * logcf(-y1 / 2#, i, 1#) - (2# / (i - 1#))) * scale2 + (scale3 * logcf(-y1 / 3#, i, 1#) + scale2 * scale2 * logcf(-y1 / 4#, i, 1#))
    'x2 = ((x1 + 2.5) * logcf(-x1 / 2#, i, 1#) - (2# / (i - 1#))) * scale2 + (scale3 * logcf(-x1 / 3#, i, 1#) + scale2 * scale2 * logcf(-x1 / 4#, i, 1#))
    y2 = ((y1 + 2.5) * logcf(-y1 / 2#, i, 1#) - (2# / (i - 1#))) * scale2 + scale3 * logcf(-y1 / 3#, i, 1#)
    x2 = ((x1 + 2.5) * logcf(-x1 / 2#, i, 1#) - (2# / (i - 1#))) * scale2 + scale3 * logcf(-x1 / 3#, i, 1#)
    If a > 0.000006 Then
       acc = y2 - x2  'This calculation is not accurate enough for b < 0 and a small - hence If b < 0 code above and derivative code below for small a
    Else
       y3 = -(y1 + a / 2#) / 2#
       x3 = -(y1 + a / 2#) / 3#
       acc = -a * (scale2 * (logcf(y3, i, 1#) + (y3 - 1.25) * (1# / (1# - y3) - i * logcf(y3, i + 1#, 1#))) - scale3 / 3# * ((1# / (1# - x3) - i * logcf(x3, i + 1#, 1#))))
    End If
    For i = UBound(coeffs2) To 0 Step -1
       acc = (a * y2 - x1 * acc)
       y2 = (coeffs2(i) - y1 * y2)
       x2 = (coeffs2(i) - x1 * x2)
    Next i
    lfbaccdif1 = s1 + (y1 * y1 * acc - a * (x2 * (x1 + y1) + Onep25Minusln2Minuseulers_const))
  Else
    lfbaccdif1 = [#VALUE!]
  End If
End Function

Private Function logdif(ByVal pr As Double, ByVal prob As Double) As Double
   Dim temp As Double
   temp = (pr - prob) / prob
   If Abs(temp) >= 0.5 Then
      logdif = Log(pr / prob)
   Else
      logdif = log0(temp)
   End If
End Function

Private Function cnormal(ByVal x As Double) As Double
'//Probability that a normal variate <= x
  Dim acc As Double, x2 As Double, d As Double, term As Double, a1 As Double, a2 As Double, b1 As Double, b2 As Double, c1 As Double, c2 As Double, c3 As Double

  If (Abs(x) < 1.5) Then
     acc = 0#
     x2 = x * x
     term = 1#
     d = 3#

     While (term > sumAcc * acc)

        d = d + 2#
        term = term * x2 / d
        acc = acc + term

     Wend

     acc = 1# + x2 / 3# * (1# + acc)
     cnormal = 0.5 + Exp(-x * x * 0.5) * x * acc * OneOverSqrTwoPi
  ElseIf (Abs(x) > 40#) Then
     If (x > 0#) Then
        cnormal = 1#
     Else
        cnormal = 0#
     End If
  Else
     x2 = x * x
     a1 = 2#
     b1 = x2 + 5#
     c2 = x2 + 9#
     a2 = a1 * c2
     b2 = b1 * c2 - 12#
     c1 = 5#
     c2 = c2 + 4#

     While ((Abs(a2 * b1 - a1 * b2) > Abs(cfVSmall * b1 * a2)))

       c3 = c1 * (c1 + 1#)
       a1 = c2 * a2 - c3 * a1
       b1 = c2 * b2 - c3 * b1
       c1 = c1 + 2#
       c2 = c2 + 4#
       c3 = c1 * (c1 + 1#)
       a2 = c2 * a1 - c3 * a2
       b2 = c2 * b1 - c3 * b2
       c1 = c1 + 2#
       c2 = c2 + 4#
       If (b2 > scalefactor) Then
         a1 = a1 * scalefactor2
         b1 = b1 * scalefactor2
         a2 = a2 * scalefactor2
         b2 = b2 * scalefactor2
       End If

     Wend

     If (x > 0#) Then
        cnormal = 1# - Exp(-x * x * 0.5) * OneOverSqrTwoPi * x / (x2 + 1# - a2 / b2)
     Else
        cnormal = -Exp(-x * x * 0.5) * OneOverSqrTwoPi * x / (x2 + 1# - a2 / b2)
     End If

  End If
End Function

Private Function invcnormal(ByVal p As Double) As Double
'//Inverse of cnormal from AS241.
'//Require p to be strictly in the range 0..1

   Dim PPND16 As Double, q As Double, r As Double
   q = p - 0.5
   If (Abs(q) <= 0.425) Then
      r = 0.180625 - q * q
      PPND16 = q * (((((((a7 * r + a6) * r + a5) * r + a4) * r + a3) * r + a2) * r + a1) * r + a0) / (((((((b7 * r + b6) * r + b5) * r + b4) * r + b3) * r + b2) * r + b1) * r + 1#)
   Else
      If (q < 0#) Then
         r = p
      Else
         r = 1# - p
      End If
      r = Sqr(-Log(r))
      If (r <= 5#) Then
        r = r - 1.6
        PPND16 = (((((((c7 * r + c6) * r + c5) * r + c4) * r + c3) * r + c2) * r + c1) * r + c0) / (((((((d7 * r + d6) * r + d5) * r + d4) * r + d3) * r + d2) * r + d1) * r + 1#)
      Else
        r = r - 5#
        PPND16 = (((((((e7 * r + e6) * r + e5) * r + e4) * r + e3) * r + e2) * r + e1) * r + e0) / (((((((f7 * r + f6) * r + f5) * r + f4) * r + f3) * r + f2) * r + f1) * r + 1#)
      End If
      If (q < 0#) Then
         PPND16 = -PPND16
      End If
   End If
   invcnormal = PPND16
End Function

Public Function pdf_lognormal(ByVal x As Double, ByVal mean As Double, ByVal sd As Double) As Double
   If (sd <= 0#) Then
      pdf_lognormal = [#VALUE!]
   Else
      pdf_lognormal = Exp(-0.5 * ((Log(x) - mean) / sd) ^ 2) / x / sd * OneOverSqrTwoPi
   End If
End Function

Public Function cdf_lognormal(ByVal x As Double, ByVal mean As Double, ByVal sd As Double) As Double
   If (sd <= 0#) Then
      cdf_lognormal = [#VALUE!]
   Else
      cdf_lognormal = cnormal((Log(x) - mean) / sd)
   End If
End Function

Public Function comp_cdf_lognormal(ByVal x As Double, ByVal mean As Double, ByVal sd As Double) As Double
   If (sd <= 0#) Then
      comp_cdf_lognormal = [#VALUE!]
   Else
      comp_cdf_lognormal = cnormal(-(Log(x) - mean) / sd)
   End If
End Function

Public Function inv_lognormal(ByVal prob As Double, ByVal mean As Double, ByVal sd As Double) As Double
   If (prob <= 0# Or prob >= 1# Or sd <= 0#) Then
      inv_lognormal = [#VALUE!]
   Else
      inv_lognormal = Exp(mean + sd * invcnormal(prob))
   End If
End Function

Public Function comp_inv_lognormal(ByVal prob As Double, ByVal mean As Double, ByVal sd As Double) As Double
   If (prob <= 0# Or prob >= 1# Or sd <= 0#) Then
      comp_inv_lognormal = [#VALUE!]
   Else
      comp_inv_lognormal = Exp(mean - sd * invcnormal(prob))
   End If
End Function

Private Function tdistexp(ByVal p As Double, ByVal q As Double, ByVal logqk2 As Double, ByVal k As Double, ByRef tdistDensity As Double) As Double
'//Special transformation of t-distribution useful for BinApprox.
'//Note approxtdistDens only used by binApprox if k > 100 or so.
   Dim sum As Double, aki As Double, ai As Double, term As Double, q1 As Double, q8 As Double
   Dim c1 As Double, c2 As Double, a1 As Double, a2 As Double, b1 As Double, b2 As Double, cadd As Double
   Dim result As Double, approxtdistDens As Double

   approxtdistDens = Exp(logqk2 + logfbit(k - 1#) - 2# * logfbit(k * 0.5 - 1#)) * OneOverSqrTwoPi

   If (k * p < 4# * q) Then
     sum = 0#
     aki = k + 1#
     ai = 3#
     term = 1#

     While (term > sumAcc * sum)

        ai = ai + 2#
        aki = aki + 2#
        term = term * aki * p / ai
        sum = sum + term

     Wend

     sum = 1# + (k + 1#) * p * (1# + sum) / 3#
     result = 0.5 - approxtdistDens * sum * Sqr(k * p)
   ElseIf approxtdistDens = 0# Then
     result = 0#
   Else
     q1 = 2# * (1# + q)
     q8 = 8# * q
     a1 = 1#
     b1 = (k - 3#) * p + 7#
     c1 = -20# * q
     a2 = (k - 5#) * p + 11#
     b2 = a2 * b1 + c1
     cadd = -30# * q
     c1 = -42# * q
     c2 = (k - 7#) * p + 15#

     While ((Abs(a2 * b1 - a1 * b2) > Abs(cfVSmall * b1 * a2)))

       a1 = c2 * a2 + c1 * a1
       b1 = c2 * b2 + c1 * b1
       c1 = c1 + cadd
       cadd = cadd - q8
       c2 = c2 + q1
       a2 = c2 * a1 + c1 * a2
       b2 = c2 * b1 + c1 * b2
       c1 = c1 + cadd
       cadd = cadd - q8
       c2 = c2 + q1
       If (Abs(b2) > scalefactor) Then
         a1 = a1 * scalefactor2
         b1 = b1 * scalefactor2
         a2 = a2 * scalefactor2
         b2 = b2 * scalefactor2
       ElseIf (Abs(b2) < scalefactor2) Then
         a1 = a1 * scalefactor
         b1 = b1 * scalefactor
         a2 = a2 * scalefactor
         b2 = b2 * scalefactor
       End If
     Wend

     result = approxtdistDens * (1# - q / ((k - 1#) * p + 3# - 6# * q * a2 / b2)) / Sqr(k * p)
   End If
   tdistDensity = approxtdistDens * Sqr(q)
   tdistexp = result
End Function

Private Function tdist(ByVal x As Double, ByVal k As Double, tdistDensity As Double) As Double
'//Probability that variate from t-distribution with k degress of freedom <= x
   Dim x2 As Double, k2 As Double, logterm As Double, a As Double, r As Double, c5 As Double

   If Abs(x) >= Min(1#, k) Then
      k2 = k / x
      x2 = x + k2
      k2 = k2 / x2
      x2 = x / x2
   Else
      x2 = x * x
      k2 = k + x2
      x2 = x2 / k2
      k2 = k / k2
   End If
   If (k > 1E+30) Then
      tdist = cnormal(x)
      tdistDensity = Exp(-x * x * 0.5) * OneOverSqrTwoPi
   Else
      a = k * 0.5
      If (k2 < cSmall) Then
        logterm = (Log(k) - 2# * Log(Abs(x)))
      ElseIf (Abs(x2) < 0.5) Then
        logterm = log0(-x2)
      Else
        logterm = Log(k2)
      End If
      If (k >= 1#) Then
         logterm = logterm * a
         If (x < 0#) Then
           tdist = tdistexp(x2, k2, logterm, k, tdistDensity)
         Else
           tdist = 1# - tdistexp(x2, k2, logterm, k, tdistDensity)
         End If
         Exit Function
      End If
      c5 = -1# / (k + 2#)
      tdistDensity = Exp((a + 0.5) * logterm + a * log1(c5) - c5 + lfbaccdif1(0.5, a - 0.5)) * Sqr(a / ((1# + a))) * OneOverSqrTwoPi
      If (k2 < cSmall) Then
        r = (a + 1#) * log1(a / 1.5) - lfbaccdif1(a, 0.5) - lngammaexpansion(a)
        r = r + a * ((a - 0.5) / 1.5 + Log1p5 + (Log(k) - 2# * Log(Abs(x))))
        r = Exp(r) * (0.25 / (a + 0.5))
        If x < 0# Then
           tdist = r
        Else
           tdist = 1# - r
        End If
      ElseIf (x < 0#) Then
        If x2 < k2 Then
          tdist = 0.5 * compbeta(x2, 0.5, a)
        Else
          tdist = 0.5 * beta(k2, a, 0.5)
        End If
      Else
        If x2 < k2 Then
          tdist = 0.5 * (1# + beta(x2, 0.5, a))
        Else
          tdist = 0.5 * (1# + compbeta(k2, a, 0.5))
        End If
      End If
   End If
End Function

Private Function BetterThanTailApprox(ByVal prob As Double, ByVal df As Double) As Boolean
If df <= 2 Then
   BetterThanTailApprox = prob > 0.25 * Exp((1# - df) * 1.78514841051368)
ElseIf df <= 5 Then
   BetterThanTailApprox = prob > 0.045 * Exp((2# - df) * 1.30400766847605)
ElseIf df <= 20 Then
   BetterThanTailApprox = prob > 0.0009 * Exp((5# - df) * 0.921034037197618)
Else
   BetterThanTailApprox = prob > 0.0000000009 * Exp((20# - df) * 0.690775527898214)
End If
End Function

Private Function invtdist(ByVal prob As Double, ByVal df As Double) As Double
'//Inverse of tdist
'//Require prob to be in the range 0..1 df should be positive
  Dim xn As Double, xn2 As Double, tp As Double, tpDif As Double, tprob As Double, a As Double, pr As Double, lpr As Double, small As Double, smalllpr As Double, tdistDensity As Double
  If prob > 0.5 Then
     pr = 1# - prob
  Else
     pr = prob
  End If
  lpr = -Log(pr)
  small = 0.00000000000001
  smalllpr = small * lpr * pr
  If pr >= 0.5 Or df >= 1# And BetterThanTailApprox(pr, df) Then
'// Will divide by 0 if tp so small that tdistDensity underflows. Not a problem if prob > cSmall
     xn = invcnormal(pr)
     xn2 = xn * xn
'//Initial approximation is given in http://digital.library.adelaide.edu.au/coll/special//fisher/281.pdf. The modified NR correction then gets it right.
     tp = (((((27# * xn2 + 339#) * xn2 + 930#) * xn2 - 1782#) * xn2 - 765#) * xn2 + 17955#) / (368640# * df)
     tp = (tp + ((((79# * xn2 + 776#) * xn2 + 1482#) * xn2 - 1920#) * xn2 - 945#) / 92160#) / df
     tp = (tp + (((3# * xn2 + 19#) * xn2 + 17#) * xn2 - 15#) / 384#) / df
     tp = (tp + ((5# * xn2 + 16) * xn2 + 3#) / 96#) / df
     tp = (tp + (xn2 + 1#) / 4#) / df
     tp = xn * (1# + tp)
     tprob = 0#
     tpDif = 1# + Abs(tp)
  ElseIf df < 1# Then
     a = df / 2#
     tp = (a + 1#) * log1(a / 1.5) - lfbaccdif1(a, 0.5) - lngammaexpansion(a)
     tp = ((a - 0.5) / 1.5 + Log1p5 + Log(df)) / 2# + (tp - Log(4# * pr * (a + 0.5))) / df
     tp = -Exp(tp)
     tprob = tdist(tp, df, tdistDensity)
     If tdistDensity < nearly_zero Then
        tpDif = 0#
     Else
        tpDif = (tprob / tdistDensity) * log0((tprob - pr) / pr)
        tp = tp - tpDif
     End If
  Else
     tp = tdist(0, df, tdistDensity) 'Marginally quicker to get tdistDensity for integral df
     tp = Exp(-Log(Sqr(df) * pr / tdistDensity) / df)
     If df >= 2 Then
        tp = -Sqr(df * (tp * tp - 1#))
     Else
        tp = -Sqr(df) * Sqr(tp - 1#) * Sqr(tp + 1#)
     End If
     tpDif = tp / df
     tpDif = -log0((0.5 - 1# / (df + 2)) / (1# + tpDif * tp)) * (tpDif + 1# / tp)
     tp = tp - tpDif
     tprob = tdist(tp, df, tdistDensity)
     If tdistDensity < nearly_zero Then
        tpDif = 0#
     Else
        tpDif = (tprob / tdistDensity) * log0((tprob - pr) / pr)
        tp = tp - tpDif
     End If
  End If
  While (Abs(tprob - pr) > smalllpr And Abs(tpDif) > small * (1# + Abs(tp)))
     tprob = tdist(tp, df, tdistDensity)
     tpDif = (tprob / tdistDensity) * log0((tprob - pr) / pr)
     tp = tp - tpDif
  Wend
  invtdist = tp
  If prob > 0.5 Then invtdist = -invtdist
End Function

Private Function poissonTerm(ByVal i As Double, ByVal n As Double, ByVal diffFromMean As Double, ByVal logAdd As Double) As Double
'//Probability that poisson variate with mean n has value i (diffFromMean = n-i)
   Dim c2 As Double, c3 As Double
   Dim logpoissonTerm As Double, c1 As Double

   If ((i <= -1#) Or (n < 0#)) Then
      If (i = 0#) Then
         poissonTerm = Exp(logAdd)
      Else
         poissonTerm = 0#
      End If
   ElseIf ((i < 0#) And (n = 0#)) Then
      poissonTerm = [#VALUE!]
   Else
     c3 = i
     c2 = c3 + 1#
     c1 = (diffFromMean - 1#) / c2

     If (c1 < minLog1Value) Then
        If (i = 0#) Then
          logpoissonTerm = -n
          poissonTerm = Exp(logpoissonTerm + logAdd)
        Else
          On Error GoTo ptiszero
          logpoissonTerm = (c3 * Log(n / c2) - (diffFromMean - 1#)) - logfbit(c3)
          poissonTerm = Exp(logpoissonTerm + logAdd) / Sqr(c2) * OneOverSqrTwoPi
          Exit Function
ptiszero: poissonTerm = 0#
          Exit Function
        End If
     Else
       logpoissonTerm = c3 * log1(c1) - c1 - logfbit(c3)
       poissonTerm = Exp(logpoissonTerm + logAdd) / Sqr(c2) * OneOverSqrTwoPi
     End If
   End If
End Function

Private Function poisson1(ByVal i As Double, ByVal n As Double, ByVal diffFromMean As Double) As Double
'//Probability that poisson variate with mean n has value <= i (diffFromMean = n-i)
'//For negative values of i (used for calculating the cumlative gamma distribution) there's a really nasty interpretation!
'//1-gamma(n,i) is calculated as poisson1(-i,n,0) since we need an accurate version of i rather than i-1.
'//Uses a simplified version of Legendre's continued fraction.
   Dim prob As Double, exact As Boolean
   If ((i >= 0#) And (n <= 0#)) Then
      exact = True
      prob = 1#
   ElseIf ((i > -1#) And (n <= 0#)) Then
      exact = True
      prob = 0#
   ElseIf ((i > -1#) And (i < 0#)) Then
      i = -i
      exact = False
      prob = poissonTerm(i, n, n - i, 0#) * i / n
      i = i - 1#
      diffFromMean = n - i
   Else
      exact = ((i <= -1#) Or (n < 0#))
      prob = poissonTerm(i, n, diffFromMean, 0#)
   End If
   If (exact Or prob = 0#) Then
      poisson1 = prob
      Exit Function
   End If

   Dim a1 As Double, a2 As Double, b1 As Double, b2 As Double, c1 As Double, c2 As Double, c3 As Double, c4 As Double, cfValue As Double
   Dim njj As Long, numb As Long
   Dim sumAlways As Long, sumFactor As Long
   sumAlways = 0
   sumFactor = 6
   a1 = 0#
   If (i > sumAlways) Then
      numb = Int(sumFactor * Exp(Log(n) / 3))
      numb = Max(0, Int(numb - diffFromMean))
      If (numb > i) Then
         numb = Int(i)
      End If
   Else
      numb = Max(0, Int(i))
   End If

   b1 = 1#
   a2 = i - numb
   b2 = diffFromMean + (numb + 1#)
   c1 = 0#
   c2 = a2
   c4 = b2
   If c2 < 0# Then
      cfValue = cfVSmall
   Else
      cfValue = cfSmall
   End If
   While ((Abs(a2 * b1 - a1 * b2) > Abs(cfValue * b1 * a2)))

       c1 = c1 + 1#
       c2 = c2 - 1#
       c3 = c1 * c2
       c4 = c4 + 2#
       a1 = c4 * a2 + c3 * a1
       b1 = c4 * b2 + c3 * b1
       c1 = c1 + 1#
       c2 = c2 - 1#
       c3 = c1 * c2
       c4 = c4 + 2#
       a2 = c4 * a1 + c3 * a2
       b2 = c4 * b1 + c3 * b2
       If (b2 > scalefactor) Then
         a1 = a1 * scalefactor2
         b1 = b1 * scalefactor2
         a2 = a2 * scalefactor2
         b2 = b2 * scalefactor2
       End If
       If c2 < 0# And cfValue > cfVSmall Then
          cfValue = cfVSmall
       End If
   Wend

   a1 = a2 / b2

   c1 = i - numb + 1#
   For njj = 1 To numb
     a1 = (1# + a1) * (c1 / n)
     c1 = c1 + 1#
   Next njj

   poisson1 = (1# + a1) * prob
End Function

Private Function poisson2(ByVal i As Double, ByVal n As Double, ByVal diffFromMean As Double) As Double
'//Probability that poisson variate with mean n has value >= i (diffFromMean = n-i)
   Dim prob As Double, exact As Boolean
   If ((i <= 0#) And (n <= 0#)) Then
      exact = True
      prob = 1#
   Else
      exact = False
      prob = poissonTerm(i, n, diffFromMean, 0#)
   End If
   If (exact Or prob = 0#) Then
      poisson2 = prob
      Exit Function
   End If

   Dim a1 As Double, a2 As Double, b1 As Double, b2 As Double, c1 As Double, c2 As Double
   Dim njj As Long, numb As Long
   Const sumFactor = 6
   numb = Int(sumFactor * Exp(Log(n) / 3))
   numb = Max(0, Int(diffFromMean + numb))

   a1 = 0#
   b1 = 1#
   a2 = n
   b2 = (numb + 1#) - diffFromMean
   c1 = 0#
   c2 = b2

   While ((Abs(a2 * b1 - a1 * b2) > Abs(cfSmall * b1 * a2)))

      c1 = c1 + n
      c2 = c2 + 1#
      a1 = c2 * a2 + c1 * a1
      b1 = c2 * b2 + c1 * b1
      c1 = c1 + n
      c2 = c2 + 1#
      a2 = c2 * a1 + c1 * a2
      b2 = c2 * b1 + c1 * b2
      If (b2 > scalefactor) Then
        a1 = a1 * scalefactor2
        b1 = b1 * scalefactor2
        a2 = a2 * scalefactor2
        b2 = b2 * scalefactor2
      End If
   Wend

   a1 = a2 / b2

   c1 = i + numb
   For njj = 1 To numb
     a1 = (1# + a1) * (n / c1)
     c1 = c1 - 1#
   Next

   poisson2 = (1# + a1) * prob

End Function

Private Function poissonApprox(ByVal j As Double, ByVal diffFromMean As Double, ByVal comp As Boolean) As Double
'//Asymptotic expansion to calculate the probability that poisson variate has value <= j (diffFromMean = mean-j). If comp then calulate 1-probability.
'//cf. http://members.aol.com/iandjmsmith/PoissonApprox.htm
Dim pt As Double, s2pt As Double, res1 As Double, res2 As Double, elfb As Double, term As Double
Dim ig2 As Double, ig3 As Double, ig4 As Double, ig5 As Double, ig6 As Double, ig7 As Double, ig8 As Double
Dim ig05 As Double, ig25 As Double, ig35 As Double, ig45 As Double, ig55 As Double, ig65 As Double, ig75 As Double

pt = -log1(diffFromMean / j)
s2pt = Sqr(2# * j * pt)

ig2 = 1# / j + pt
term = pt * pt * 0.5
ig3 = ig2 / j + term
term = term * pt / 3#
ig4 = ig3 / j + term
term = term * pt / 4#
ig5 = ig4 / j + term
term = term * pt / 5#
ig6 = ig5 / j + term
term = term * pt / 6#
ig7 = ig6 / j + term
term = term * pt / 7#
ig8 = ig7 / j + term

ig05 = cnormal(-s2pt)
term = pt * twoThirds
ig25 = 1# / j + term
term = term * pt * twoFifths
ig35 = ig25 / j + term
term = term * pt * twoSevenths
ig45 = ig35 / j + term
term = term * pt * twoNinths
ig55 = ig45 / j + term
term = term * pt * twoElevenths
ig65 = ig55 / j + term
term = term * pt * twoThirteenths
ig75 = ig65 / j + term

elfb = ((((((coef75 / j + coef65) / j + coef55) / j + coef45) / j + coef35) / j + coef25) / j + coef15) + j
res1 = (((((((ig8 * coef8 + ig7 * coef7) + ig6 * coef6) + ig5 * coef5) + ig4 * coef4) + ig3 * coef3) + ig2 * coef2) + coef1) * Sqr(j)
res2 = ((((((ig75 * coef75 + ig65 * coef65) + ig55 * coef55) + ig45 * coef45) + ig35 * coef35) + ig25 * coef25) + coef15) * s2pt

If (comp) Then
   If (diffFromMean < 0#) Then
      poissonApprox = ig05 - (res1 - res2) * Exp(-j * pt) * OneOverSqrTwoPi / elfb
   Else
      poissonApprox = (1# - ig05) - (res1 + res2) * Exp(-j * pt) * OneOverSqrTwoPi / elfb
   End If
ElseIf (diffFromMean < 0#) Then
   poissonApprox = (1# - ig05) + (res1 - res2) * Exp(-j * pt) * OneOverSqrTwoPi / elfb
Else
   poissonApprox = ig05 + (res1 + res2) * Exp(-j * pt) * OneOverSqrTwoPi / elfb
End If
End Function

Private Function cpoisson(ByVal k As Double, ByVal lambda As Double, ByVal dfm As Double) As Double
'//Probability that poisson variate with mean lambda has value <= k (diffFromMean = lambda-k) calculated by various methods.
   If ((k >= 21#) And (Abs(dfm) < (0.3 * k))) Then
      cpoisson = poissonApprox(k, dfm, False)
   ElseIf ((lambda > k) And (lambda >= 1#)) Then
      cpoisson = poisson1(k, lambda, dfm)
   Else
      cpoisson = 1# - poisson2(k + 1#, lambda, dfm - 1#)
   End If
End Function

Private Function comppoisson(ByVal k As Double, ByVal lambda As Double, ByVal dfm As Double) As Double
'//Probability that poisson variate with mean lambda has value > k (diffFromMean = lambda-k) calculated by various methods.
   If ((k >= 21#) And (Abs(dfm) < (0.3 * k))) Then
      comppoisson = poissonApprox(k, dfm, True)
   ElseIf ((lambda > k) And (lambda >= 1#)) Then
      comppoisson = 1# - poisson1(k, lambda, dfm)
   Else
      comppoisson = poisson2(k + 1#, lambda, dfm - 1#)
   End If
End Function

Private Function invpoisson(ByVal k As Double, ByVal prob As Double) As Double
'//Inverse of poisson. Calculates mean such that poisson(k,mean,mean-k)=prob.
'//Require prob to be in the range 0..1, k should be -1/2 or non-negative
   If (k = 0#) Then
      invpoisson = -Log(prob + 9.99988867182683E-321)
   ElseIf (prob > 0.5) Then
      invpoisson = invcomppoisson(k, 1# - prob)
   Else '/*if (k > 0#)*/ then
      Dim temp2 As Double, xp As Double, dfm As Double, q As Double, qdif As Double, lpr As Double, small As Double, smalllpr As Double
      lpr = -Log(prob)
      small = 0.00000000000001
      smalllpr = small * lpr * prob
      xp = invcnormal(prob)
      dfm = xp * (0.5 * xp - Sqr(k + (0.5 * xp) ^ 2))
      q = -1#
      qdif = -dfm
      If Abs(qdif) < 1# Then
         qdif = 1#
      ElseIf (k > 1E+50) Then
         invpoisson = k
         Exit Function
      End If
      While ((Abs(q - prob) > smalllpr) And (Abs(qdif) > (1# + Abs(dfm)) * small))
         q = cpoisson(k, k + dfm, dfm)
         If (q = 0#) Then
             qdif = qdif / 2#
             dfm = dfm + qdif
             q = -1#
         Else
            temp2 = poissonTerm(k, k + dfm, dfm, 0#)
            If (temp2 = 0#) Then
               qdif = qdif / 2#
               dfm = dfm + qdif
               q = -1#
            Else
               qdif = -2# * q * logdif(q, prob) / (1# + Sqr(Log(prob) / Log(q))) / temp2
               If (qdif > k + dfm) Then
                  qdif = dfm / 2#
                  dfm = dfm - qdif
                  q = -1#
               Else
                  dfm = dfm - qdif
               End If
            End If
         End If
      Wend
      invpoisson = k + dfm
   End If
End Function

Private Function invcomppoisson(ByVal k As Double, ByVal prob As Double) As Double
'//Inverse of comppoisson. Calculates mean such that comppoisson(k,mean,mean-k)=prob.
'//Require prob to be in the range 0..1, k should be -1/2 or non-negative
   If (prob > 0.5) Then
      invcomppoisson = invpoisson(k, 1# - prob)
   ElseIf (k = 0#) Then
      invcomppoisson = -log0(-prob)
   Else '/*if (k > 0#)*/ then
      Dim temp2 As Double, xp As Double, dfm As Double, q As Double, qdif As Double, lambda As Double, qdifset As Boolean, lpr As Double, small As Double, smalllpr As Double
      lpr = -Log(prob)
      small = 0.00000000000001
      smalllpr = small * lpr * prob
      xp = invcnormal(prob)
      dfm = xp * (0.5 * xp + Sqr(k + (0.5 * xp) ^ 2))
      lambda = k + dfm
      If ((lambda < 1#) And (k < 40#)) Then
         lambda = Exp(Log(prob / poissonTerm(k + 1#, 1#, -k, 0#)) / (k + 1#))
         dfm = lambda - k
      ElseIf (k > 1E+50) Then
         invcomppoisson = lambda
         Exit Function
      End If
      q = -1#
      qdif = lambda
      qdifset = False
      While ((Abs(q - prob) > smalllpr) And (Abs(qdif) > Min(lambda, Abs(dfm)) * small))
         q = comppoisson(k, lambda, dfm)
         If (q = 0#) Then
            If qdifset Then
               qdif = qdif / 2#
               dfm = dfm + qdif
               lambda = lambda + qdif
            Else
               lambda = 2# * lambda
               qdif = qdif * 2#
               dfm = lambda - k
            End If
            q = -1#
         Else
            temp2 = poissonTerm(k, lambda, dfm, 0#)
            If (temp2 = 0#) Then
               If qdifset Then
                  qdif = qdif / 2#
                  dfm = dfm + qdif
                  lambda = lambda + qdif
               Else
                  lambda = 2# * lambda
                  qdif = qdif * 2#
                  dfm = lambda - k
               End If
               q = -1#
            Else
               qdif = 2# * q * logdif(q, prob) / (1# + Sqr(Log(prob) / Log(q))) / temp2
               If (qdif > lambda) Then
                  lambda = lambda / 10#
                  qdif = dfm
                  dfm = lambda - k
                  qdif = qdif - dfm
                  q = -1#
               Else
                  lambda = lambda - qdif
                  dfm = dfm - qdif
               End If
               qdifset = True
            End If
         End If
         If (Abs(dfm) > lambda) Then
            dfm = lambda - k
         Else
            lambda = k + dfm
         End If
      Wend
      invcomppoisson = lambda
   End If
End Function

Private Function binomialTerm(ByVal i As Double, ByVal j As Double, ByVal p As Double, ByVal q As Double, ByVal diffFromMean As Double, ByVal logAdd As Double) As Double
'//Probability that binomial variate with sample size i+j and event prob p (=1-q) has value i (diffFromMean = (i+j)*p-i)
   Dim c1 As Double, c2 As Double, c3 As Double
   Dim c4 As Double, c5 As Double, c6 As Double, ps As Double, logbinomialTerm As Double, dfm As Double
   If ((i = 0#) And (j <= 0#)) Then
      binomialTerm = Exp(logAdd)
   ElseIf ((i <= -1#) Or (j < 0#)) Then
      binomialTerm = 0#
   Else
      If (p < q) Then
         c2 = i
         c3 = j
         ps = p
         dfm = diffFromMean
      Else
         c3 = i
         c2 = j
         ps = q
         dfm = -diffFromMean
      End If

      c5 = (dfm - (1# - ps)) / (c2 + 1#)
      c6 = -(dfm + ps) / (c3 + 1#)

      If (c5 < minLog1Value) Then
         If (c2 = 0#) Then
            logbinomialTerm = c3 * log0(-ps)
            binomialTerm = Exp(logbinomialTerm + logAdd)
         ElseIf ((ps = 0#) And (c2 > 0#)) Then
            binomialTerm = 0#
         Else
            c1 = (i + 1#) + j
            'c4 = logfbit(i + j) - logfbit(i) - logfbit(j)
            'logbinomialTerm = c4 + c2 * (Log((ps * c1) / (c2 + 1#)) - c5) - c5 + c3 * log1(c6) - c6
            c4 = lfbaccdif1(j, i) + logfbit(j)
            logbinomialTerm = c2 * (Log((ps * c1) / (c2 + 1#)) - c5) - c5 + c3 * log1(c6) - c6 - c4
            binomialTerm = Exp(logbinomialTerm + logAdd) * Sqr(c1 / ((c2 + 1#) * (c3 + 1#))) * OneOverSqrTwoPi
         End If
      Else
         'c4 = logfbit(i + j) - logfbit(i) - logfbit(j)
         'logbinomialTerm = c4 + (c2 * log1(c5) - c5) + (c3 * log1(c6) - c6)
         c4 = lfbaccdif1(j, i) + logfbit(j)
         logbinomialTerm = (c2 * log1(c5) - c5) + (c3 * log1(c6) - c6) - c4
         binomialTerm = Exp(logbinomialTerm + logAdd) * Sqr((1# + j / (i + 1#)) / (j + 1#)) * OneOverSqrTwoPi
      End If
   End If
End Function

Private Function binomialcf(ByVal ii As Double, ByVal jj As Double, ByVal pp As Double, ByVal qq As Double, ByVal diffFromMean As Double, ByVal comp As Boolean) As Double
'//Probability that binomial variate with sample size ii+jj and event prob pp (=1-qq) has value <=i (diffFromMean = (ii+jj)*pp-ii). If comp the returns 1 - probability.
Dim prob As Double, p As Double, q As Double, a1 As Double, a2 As Double, b1 As Double, b2 As Double
Dim c1 As Double, c2 As Double, c3 As Double, c4 As Double, n1 As Double, q1 As Double, dfm As Double
Dim i As Double, j As Double, ni As Double, nj As Double, numb As Double, ip1 As Double, cfValue As Double
Dim swapped As Boolean, exact As Boolean

  If ((ii > -1#) And (ii < 0#)) Then
     ip1 = -ii
     ii = ip1 - 1#
  Else
     ip1 = ii + 1#
  End If
  n1 = (ii + 3#) + jj
  If ii < 0# Then
     cfValue = cfVSmall
     swapped = False
  ElseIf pp > qq Then
     cfValue = cfSmall
     swapped = n1 * qq >= jj + 1#
  Else
     cfValue = cfSmall
     swapped = n1 * pp <= ii + 2#
  End If
  If Not swapped Then
    i = ii
    j = jj
    p = pp
    q = qq
    dfm = diffFromMean
  Else
    j = ip1
    ip1 = jj
    i = jj - 1#
    p = qq
    q = pp
    dfm = 1# - diffFromMean
  End If
  If ((i > -1#) And ((j <= 0#) Or (p = 0#))) Then
     exact = True
     prob = 1#
  ElseIf ((i > -1#) And (i < 0#) Or (i = -1#) And (ip1 > 0#)) Then
     exact = False
     prob = binomialTerm(ip1, j, p, q, (ip1 + j) * p - ip1, 0#) * ip1 / ((ip1 + j) * p)
     dfm = (i + j) * p - i
  Else
     exact = ((i = 0#) And (j <= 0#)) Or ((i <= -1#) Or (j < 0#))
     prob = binomialTerm(i, j, p, q, dfm, 0#)
  End If
  If (exact) Or (prob = 0#) Then
     If (swapped = comp) Then
        binomialcf = prob
     Else
        binomialcf = 1# - prob
     End If
     Exit Function
  End If

  Dim sumAlways As Long, sumFactor As Long
  sumAlways = 0
  sumFactor = 6
  a1 = 0#
  If (i > sumAlways) Then
     numb = Int(sumFactor * Sqr(p + 0.5) * Exp(Log(n1 * p * q) / 3))
     numb = Int(numb - dfm)
     If (numb > i) Then
        numb = Int(i)
     End If
  Else
     numb = Int(i)
  End If
  If (numb < 0#) Then
     numb = 0#
  End If

  b1 = 1#
  q1 = q + 1#
  a2 = (i - numb) * q
  b2 = dfm + numb + 1#
  c1 = 0#
  c2 = a2
  c4 = b2
  While ((Abs(a2 * b1 - a1 * b2) > Abs(cfValue * b1 * a2)))

    c1 = c1 + 1#
    c2 = c2 - q
    c3 = c1 * c2
    c4 = c4 + q1
    a1 = c4 * a2 + c3 * a1
    b1 = c4 * b2 + c3 * b1
    c1 = c1 + 1#
    c2 = c2 - q
    c3 = c1 * c2
    c4 = c4 + q1
    a2 = c4 * a1 + c3 * a2
    b2 = c4 * b1 + c3 * b2
    If (Abs(b2) > scalefactor) Then
      a1 = a1 * scalefactor2
      b1 = b1 * scalefactor2
      a2 = a2 * scalefactor2
      b2 = b2 * scalefactor2
    ElseIf (Abs(b2) < scalefactor2) Then
      a1 = a1 * scalefactor
      b1 = b1 * scalefactor
      a2 = a2 * scalefactor
      b2 = b2 * scalefactor
    End If
    If c2 < 0# And cfValue > cfVSmall Then
       cfValue = cfVSmall
    End If
  Wend
  a1 = a2 / b2

  ni = (i - numb + 1#) * q
  nj = (j + numb) * p
  While (numb > 0#)
     a1 = (1# + a1) * (ni / nj)
     ni = ni + q
     nj = nj - p
     numb = numb - 1#
  Wend

  a1 = (1# + a1) * prob
  If (swapped = comp) Then
     binomialcf = a1
  Else
     binomialcf = 1# - a1
  End If

End Function

Private Function binApprox(ByVal a As Double, ByVal b As Double, ByVal diffFromMean As Double, ByVal comp As Boolean) As Double
'//Asymptotic expansion to calculate the probability that binomial variate has value <= a (diffFromMean = (a+b)*p-a). If comp then calulate 1-probability.
'//cf. http://members.aol.com/iandjmsmith/BinomialApprox.htm
Dim n As Double, n1 As Double
Dim pq1 As Double, mfac As Double, res As Double, tp As Double, lval As Double, lvv As Double, temp As Double
Dim ib05 As Double, ib15 As Double, ib25 As Double, ib35 As Double, ib45 As Double, ib55 As Double, ib65 As Double
Dim ib2 As Double, ib3 As Double, ib4 As Double, ib5 As Double, ib6 As Double, ib7 As Double
Dim elfb As Double, coef15 As Double, coef25 As Double, coef35 As Double, coef45 As Double, coef55 As Double, coef65 As Double
Dim coef2 As Double, coef3 As Double, coef4 As Double, coef5 As Double, coef6 As Double, coef7 As Double
Dim tdistDensity As Double, approxtdistDens As Double

n = a + b
n1 = n + 1#
lvv = (b + diffFromMean) / n1 - diffFromMean
lval = (a * log1(lvv / a) + b * log1(-lvv / b)) / n
tp = -expm1(lval)

pq1 = (a / n) * (b / n)

coef15 = (-17# * pq1 + 2#) / 24#
coef25 = ((-503# * pq1 + 76#) * pq1 + 4#) / 1152#
coef35 = (((-315733# * pq1 + 53310#) * pq1 + 8196#) * pq1 - 1112#) / 414720#
coef45 = (4059192# + pq1 * (15386296# - 85262251# * pq1))
coef45 = (-9136# + pq1 * (-697376 + pq1 * coef45)) / 39813120#
coef55 = (3904584040# + pq1 * (10438368262# - 55253161559# * pq1))
coef55 = (5244128# + pq1 * (-43679536# + pq1 * (-703410640# + pq1 * coef55))) / 6688604160#
coef65 = (-3242780782432# + pq1 * (18320560326516# + pq1 * (38020748623980# - 194479285104469# * pq1)))
coef65 = (335796416# + pq1 * (61701376704# + pq1 * (-433635420336# + pq1 * coef65))) / 4815794995200#
elfb = (((((coef65 / ((n + 6.5) * pq1) + coef55) / ((n + 5.5) * pq1) + coef45) / ((n + 4.5) * pq1) + coef35) / ((n + 3.5) * pq1) + coef25) / ((n + 2.5) * pq1) + coef15) / ((n + 1.5) * pq1) + 1#

coef2 = (-pq1 - 2#) / 135#
coef3 = ((-44# * pq1 - 86#) * pq1 + 4#) / 2835#
coef4 = (((-404# * pq1 - 786#) * pq1 + 48#) * pq1 + 8#) / 8505#
coef5 = (((((-2421272# * pq1 - 4721524#) * pq1 + 302244#) * pq1) + 118160#) * pq1 - 4496#) / 12629925#
coef6 = ((((((-473759128# * pq1 - 928767700#) * pq1 + 57300188#) * pq1) + 38704888#) * pq1 - 1870064#) * pq1 - 167072#) / 492567075#
coef7 = (((((((-8530742848# * pq1 - 16836643200#) * pq1 + 954602040#) * pq1) + 990295352#) * pq1 - 44963088#) * pq1 - 11596512#) * pq1 + 349376#) / 1477701225#

ib05 = tdistexp(tp, 1# - tp, n1 * lval, 2# * n1, tdistDensity)
mfac = n1 * tp
ib15 = Sqr(2# * mfac)

If (mfac > 1E+50) Then
   ib2 = (1# + mfac) / (n + 2#)
   mfac = mfac * tp / 2#
   ib3 = (ib2 + mfac) / (n + 3#)
   mfac = mfac * tp / 3#
   ib4 = (ib3 + mfac) / (n + 4#)
   mfac = mfac * tp / 4#
   ib5 = (ib4 + mfac) / (n + 5#)
   mfac = mfac * tp / 5#
   ib6 = (ib5 + mfac) / (n + 6#)
   mfac = mfac * tp / 6#
   ib7 = (ib6 + mfac) / (n + 7#)
   res = (ib2 * coef2 + (ib3 * coef3 + (ib4 * coef4 + (ib5 * coef5 + (ib6 * coef6 + ib7 * coef7 / pq1) / pq1) / pq1) / pq1) / pq1) / pq1

   mfac = (n + 1.5) * tp * twoThirds
   ib25 = (1# + mfac) / (n + 2.5)
   mfac = mfac * tp * twoFifths
   ib35 = (ib25 + mfac) / (n + 3.5)
   mfac = mfac * tp * twoSevenths
   ib45 = (ib35 + mfac) / (n + 4.5)
   mfac = mfac * tp * twoNinths
   ib55 = (ib45 + mfac) / (n + 5.5)
   mfac = mfac * tp * twoElevenths
   ib65 = (ib55 + mfac) / (n + 6.5)
   temp = (((((coef65 * ib65 / pq1 + coef55 * ib55) / pq1 + coef45 * ib45) / pq1 + coef35 * ib35) / pq1 + coef25 * ib25) / pq1 + coef15)
Else
   ib2 = 1# + mfac
   mfac = mfac * (n + 2#) * tp / 2#
   ib3 = ib2 + mfac
   mfac = mfac * (n + 3#) * tp / 3#
   ib4 = ib3 + mfac
   mfac = mfac * (n + 4#) * tp / 4#
   ib5 = ib4 + mfac
   mfac = mfac * (n + 5#) * tp / 5#
   ib6 = ib5 + mfac
   mfac = mfac * (n + 6#) * tp / 6#
   ib7 = ib6 + mfac
   res = (ib2 * coef2 + (ib3 * coef3 + (ib4 * coef4 + (ib5 * coef5 + (ib6 * coef6 + ib7 * coef7 / ((n + 7#) * pq1)) / ((n + 6#) * pq1)) / ((n + 5#) * pq1)) / ((n + 4#) * pq1)) / ((n + 3#) * pq1)) / ((n + 2#) * pq1)

   mfac = (n + 1.5) * tp * twoThirds
   ib25 = 1# + mfac
   mfac = mfac * (n + 2.5) * tp * twoFifths
   ib35 = ib25 + mfac
   mfac = mfac * (n + 3.5) * tp * twoSevenths
   ib45 = ib35 + mfac
   mfac = mfac * (n + 4.5) * tp * twoNinths
   ib55 = ib45 + mfac
   mfac = mfac * (n + 5.5) * tp * twoElevenths
   ib65 = ib55 + mfac
   temp = (((((coef65 * ib65 / ((n + 6.5) * pq1) + coef55 * ib55) / ((n + 5.5) * pq1) + coef45 * ib45) / ((n + 4.5) * pq1) + coef35 * ib35) / ((n + 3.5) * pq1) + coef25 * ib25) / ((n + 2.5) * pq1) + coef15)
End If

approxtdistDens = tdistDensity / Sqr(1# - tp)
temp = ib15 * temp / ((n + 1.5) * pq1)
res = (oneThird + res) * 2# * (a - b) / (n * Sqr(n1 * pq1))
If (comp) Then
   If (lvv > 0#) Then
      binApprox = ib05 - (res - temp) * approxtdistDens / elfb
   Else
      binApprox = (1# - ib05) - (res + temp) * approxtdistDens / elfb
   End If
ElseIf (lvv > 0#) Then
   binApprox = (1# - ib05) + (res - temp) * approxtdistDens / elfb
Else
   binApprox = ib05 + (res + temp) * approxtdistDens / elfb
End If
End Function

Private Function binomial(ByVal ii As Double, ByVal jj As Double, ByVal pp As Double, ByVal qq As Double, ByVal diffFromMean As Double) As Double
'//Probability that binomial variate with sample size ii+jj and event prob pp (=1-qq) has value <=i (diffFromMean = (ii+jj)*pp-ii).
   Dim mij As Double
   mij = Min(ii, jj)
   If ((mij > 50#) And (Abs(diffFromMean) < (0.1 * mij))) Then
      binomial = binApprox(jj - 1#, ii, diffFromMean, False)
   Else
      binomial = binomialcf(ii, jj, pp, qq, diffFromMean, False)
   End If
End Function

Private Function compbinomial(ByVal ii As Double, ByVal jj As Double, ByVal pp As Double, ByVal qq As Double, ByVal diffFromMean As Double) As Double
'//Probability that binomial variate with sample size ii+jj and event prob pp (=1-qq) has value >i (diffFromMean = (ii+jj)*pp-ii).
   Dim mij As Double
   mij = Min(ii, jj)
   If ((mij > 50#) And (Abs(diffFromMean) < (0.1 * mij))) Then
       compbinomial = binApprox(jj - 1#, ii, diffFromMean, True)
   Else
       compbinomial = binomialcf(ii, jj, pp, qq, diffFromMean, True)
   End If
End Function

Private Function invbinom(ByVal k As Double, ByVal m As Double, ByVal prob As Double, ByRef oneMinusP As Double) As Double
'//Inverse of binomial. Delivers event probability p (q held in oneMinusP in case required) so that binomial(k,m,p,oneMinusp,dfm) = prob.
'//Note that dfm is calculated accurately but never made available outside of this routine.
'//Require prob to be in the range 0..1, m should be positive and k should be >= 0
   Dim temp1 As Double, temp2 As Double
   If (prob > 0.5) Then
      temp2 = invcompbinom(k, m, 1# - prob, oneMinusP)
   Else
      temp1 = invcompbinom(m - 1#, k + 1#, prob, oneMinusP)
      temp2 = oneMinusP
      oneMinusP = temp1
   End If
   invbinom = temp2
End Function

Private Function invcompbinom(ByVal k As Double, ByVal m As Double, ByVal prob As Double, ByRef oneMinusP As Double) As Double
'//Inverse of compbinomial. Delivers event probability p (q held in oneMinusP in case required) so that compbinomial(k,m,p,oneMinusp,dfm) = prob.
'//Note that dfm is calculated accurately but never made available outside of this routine.
'//Require prob to be in the range 0..1, m should be positive and k should be >= -0.5
Dim xp As Double, xp2 As Double, dfm As Double, n As Double, p As Double, q As Double, pr As Double, dif As Double, temp As Double, temp2 As Double, result As Double, lpr As Double, small As Double, smalllpr As Double, nminpq As Double
   result = -1#
   n = k + m
   If (prob > 0.5) Then
      result = invbinom(k, m, 1# - prob, oneMinusP)
   ElseIf (k = 0#) Then
      result = log0(-prob) / n
      If (Abs(result) < 1#) Then
        result = -expm1(result)
        oneMinusP = 1# - result
      Else
        oneMinusP = Exp(result)
        result = 1# - oneMinusP
      End If
   ElseIf (m = 1#) Then
      result = Log(prob) / n
      If (Abs(result) < 1#) Then
        oneMinusP = -expm1(result)
        result = 1# - oneMinusP
      Else
        result = Exp(result)
        oneMinusP = 1# - result
      End If
   Else
      pr = -1#
      xp = invcnormal(prob)
      xp2 = xp * xp
      temp = 2# * xp * Sqr(k * (m / n) + xp2 / 4#)
      xp2 = xp2 / n
      dfm = (xp2 * (m - k) + temp) / (2# * (1# + xp2))
      If (k + dfm < 0#) Then
         dfm = -k
      End If
      q = (m - dfm) / n
      p = (k + dfm) / n
      dif = -dfm / n
      If (dif = 0#) Then
         dif = 1#
      ElseIf Min(k, m) > 1E+50 Then
         oneMinusP = q
         invcompbinom = p
         Exit Function
      End If
      lpr = -Log(prob)
      small = 0.00000000000004
      smalllpr = small * lpr * prob
      nminpq = n * Min(p, q)
      While ((Abs(pr - prob) > smalllpr) And (n * Abs(dif) > Min(Abs(dfm), nminpq) * small))
         pr = compbinomial(k, m, p, q, dfm)
         If (pr < nearly_zero) Then '/*Should not be happenning often */
            dif = dif / 2#
            dfm = dfm + n * dif
            p = p + dif
            q = q - dif
            pr = -1#
         Else
            temp2 = binomialTerm(k, m, p, q, dfm, 0#) * m / q
            If (temp2 < nearly_zero) Then '/*Should not be happenning often */
               dif = dif / 2#
               dfm = dfm + n * dif
               p = p + dif
               q = q - dif
               pr = -1#
            Else
               dif = 2# * pr * logdif(pr, prob) / (1# + Sqr(Log(prob) / Log(pr))) / temp2
               If (q + dif <= 0#) Then '/*not v. good */
                  dif = -0.9999 * q
                  dfm = dfm - n * dif
                  p = p - dif
                  q = q + dif
                  pr = -1#
               ElseIf (p - dif <= 0#) Then '/*v. good */
                  temp = Exp(Log(prob / pr) / (k + 1#))
                  dif = p
                  p = temp * p
                  dif = p - dif
                  dfm = n * p - k
                  q = 1# - p
                  pr = -1#
               Else
                  dfm = dfm - n * dif
                  p = p - dif
                  q = q + dif
               End If
            End If
         End If
      Wend
      result = p
      oneMinusP = q
   End If
   invcompbinom = result
End Function

Private Function abMinuscd(ByVal a As Double, ByVal b As Double, ByVal c As Double, ByVal d As Double) As Double
   Dim a1 As Double, b1 As Double, c1 As Double, d1 As Double, a2 As Double, b2 As Double, c2 As Double, d2 As Double, r1 As Double, r2 As Double, r2a As Double, r3 As Double
   a2 = Int(a / twoTo27) * twoTo27
   a1 = a - a2
   b2 = Int(b / twoTo27) * twoTo27
   b1 = b - b2
   c2 = Int(c / twoTo27) * twoTo27
   c1 = c - c2
   d2 = Int(d / twoTo27) * twoTo27
   d1 = d - d2
   r1 = a1 * b1 - c1 * d1
   r2 = (a2 * b1 - c1 * d2)
   r2a = (a1 * b2 - c2 * d1)
   r3 = a2 * b2 - c2 * d2
   If (r2a < 0#) = (r2 < 0#) Then
      abMinuscd = (r3 + 2# * r2a) + ((r2 - r2a) + r1)
   Else
      abMinuscd = r3 + ((r2a + r2) + r1)
   End If
End Function

Private Function aTimes2Powerb(ByVal a As Double, ByVal b As Integer) As Double
   If b > 709 Then
      a = (a * scalefactor) * scalefactor
      b = b - 512
   ElseIf b < -709 Then
      a = (a * scalefactor2) * scalefactor2
      b = b + 512
   End If
   aTimes2Powerb = a * (2#) ^ b
End Function

Private Function GeneralabMinuscd(ByVal a As Double, ByVal b As Double, ByVal c As Double, ByVal d As Double) As Double
   Dim s As Double, ca As Double, cb As Double, cc As Double, cd As Double
   Dim l2 As Integer, pa As Integer, pb As Integer, pc As Integer, pd As Integer
   s = a * b - c * d
   If a <= 0# Or b <= 0# Or c <= 0# Or d <= 0# Then
      GeneralabMinuscd = s
      Exit Function
   ElseIf s < 0# Then
      GeneralabMinuscd = -GeneralabMinuscd(c, d, a, b)
      Exit Function
   End If
   l2 = Int(Log(a) / Log(2#))
   pa = 51 - l2
   ca = aTimes2Powerb(a, pa)
   l2 = Int(Log(b) / Log(2#))
   pb = 51 - l2
   cb = aTimes2Powerb(b, pb)
   l2 = Int(Log(c) / Log(2#))
   pc = 51 - l2
   cc = aTimes2Powerb(c, pc)
   pd = pa + pb - pc
   cd = aTimes2Powerb(d, pd)
   GeneralabMinuscd = aTimes2Powerb(abMinuscd(ca, cb, cc, cd), -(pa + pb))
End Function

Private Function hypergeometricTerm(ByVal ai As Double, ByVal aji As Double, ByVal aki As Double, ByVal amkji As Double) As Double
'// Probability that hypergeometric variate from a population with total type Is of aki+ai, total type IIs of amkji+aji, has ai type Is and aji type IIs selected.
   Dim aj As Double, am As Double, ak As Double, amj As Double, amk As Double
   Dim cjkmi As Double, ai1 As Double, aj1 As Double, ak1 As Double, am1 As Double, aki1 As Double, aji1 As Double, amk1 As Double, amj1 As Double, amkji1 As Double
   Dim c1 As Double, c3 As Double, c4 As Double, c5 As Double, loghypergeometricTerm As Double

   ak = aki + ai
   amk = amkji + aji
   aj = aji + ai
   am = amk + ak
   amj = amkji + aki
   If (am > max_discrete) Then
      hypergeometricTerm = [#VALUE!]
      Exit Function
   End If
   If ((ai = 0#) And ((aji <= 0#) Or (aki <= 0#) Or (amj < 0#) Or (amk < 0#))) Then
      hypergeometricTerm = 1#
   ElseIf ((ai > 0#) And (Min(aki, aji) = 0#) And (Max(amj, amk) = 0#)) Then
      hypergeometricTerm = 1#
   ElseIf ((ai >= 0#) And (amkji > -1#) And (aki > -1#) And (aji >= 0#)) Then
     'c1 = logfbit(amkji) + logfbit(aki) + logfbit(aji) + logfbit(am) + logfbit(ai)
     'c1 = logfbit(amk) + logfbit(ak) + logfbit(aj) + logfbit(amj) - c1
     c1 = lfbaccdif1(ak, amk) - lfbaccdif1(ai, aki) - lfbaccdif1(ai, aji) - lfbaccdif1(aki, amkji) - logfbit(ai)
     ai1 = ai + 1#
     aj1 = aj + 1#
     ak1 = ak + 1#
     am1 = am + 1#
     aki1 = aki + 1#
     aji1 = aji + 1#
     amk1 = amk + 1#
     amj1 = amj + 1#
     amkji1 = amkji + 1#
     cjkmi = GeneralabMinuscd(aji, aki, ai, amkji)
     c5 = (cjkmi - ai) / (amkji1 * am1)
     If (c5 < minLog1Value) Then
        c3 = amkji * (Log((amj1 * amk1) / (amkji1 * am1)) - c5) - c5
     Else
        c3 = amkji * log1(c5) - c5
     End If

     c5 = (-cjkmi - aji) / (aki1 * am1)
     If (c5 < minLog1Value) Then
        c4 = aki * (Log((ak1 * amj1) / (aki1 * am1)) - c5) - c5
     Else
        c4 = aki * log1(c5) - c5
     End If

     c3 = c3 + c4
     c5 = (-cjkmi - aki) / (aji1 * am1)
     If (c5 < minLog1Value) Then
        c4 = aji * (Log((aj1 * amk1) / (aji1 * am1)) - c5) - c5
     Else
        c4 = aji * log1(c5) - c5
     End If

     c3 = c3 + c4
     c5 = (cjkmi - amkji) / (ai1 * am1)
     If (c5 < minLog1Value) Then
        c4 = ai * (Log((aj1 * ak1) / (ai1 * am1)) - c5) - c5
     Else
        c4 = ai * log1(c5) - c5
     End If

     c3 = c3 + c4
     loghypergeometricTerm = (c1 + 1# / am1) + c3

     hypergeometricTerm = Exp(loghypergeometricTerm) * Sqr((amk1 * ak1) * (aj1 * amj1) / ((amkji1 * aki1 * aji1) * (am1 * ai1))) * OneOverSqrTwoPi
   Else
     hypergeometricTerm = 0#
   End If

End Function

Private Function hypergeometric(ByVal ai As Double, ByVal aji As Double, ByVal aki As Double, ByVal amkji As Double, ByVal comp As Boolean, ByRef ha1 As Double, ByRef hprob As Double, ByRef hswap As Boolean) As Double
'// Probability that hypergeometric variate from a population with total type Is of aki+ai, total type IIs of amkji+aji, has up to ai type Is selected in a sample of size aji+ai.
     Dim prob As Double
     Dim a1 As Double, a2 As Double, b1 As Double, b2 As Double, an As Double, bn As Double, bnAdd As Double, s As Double
     Dim c1 As Double, c2 As Double, c3 As Double, c4 As Double
     Dim i As Double, ji As Double, ki As Double, mkji As Double, njj As Double, numb As Double, maxSums As Double, swapped As Boolean
     Dim ip1 As Double, must_do_cf As Boolean, allIntegral As Boolean, exact As Boolean
     If (amkji > -1#) And (amkji < 0#) Then
        ip1 = -amkji
        mkji = ip1 - 1#
        allIntegral = False
     Else
        ip1 = amkji + 1#
        mkji = amkji
        allIntegral = ai = Int(ai) And aji = Int(aji) And aki = Int(aki) And mkji = Int(mkji)
     End If

     If allIntegral Then
        swapped = (ai + 0.5) * (mkji + 0.5) >= (aki - 0.5) * (aji - 0.5)
     ElseIf ai < 100# And ai = Int(ai) Or mkji < 0# Then
        If comp Then
           swapped = (ai + 0.5) * (mkji + 0.5) >= aki * aji
        Else
           swapped = (ai + 0.5) * (mkji + 0.5) >= aki * aji + 1000#
        End If
     ElseIf ai < 1# Then
        swapped = (ai + 0.5) * (mkji + 0.5) >= aki * aji
     ElseIf aji < 1# Or aki < 1# Or (ai < 1# And ai > 0#) Then
        swapped = False
     Else
        swapped = (ai + 0.5) * (mkji + 0.5) >= (aki - 0.5) * (aji - 0.5)
     End If
     If Not swapped Then
       i = ai
       ji = aji
       ki = aki
     Else
       i = aji - 1#
       ji = ai + 1#
       ki = ip1
       ip1 = aki
       mkji = aki - 1#
     End If
     c2 = ji + i
     c4 = mkji + ki + c2
     If (c4 > max_discrete) Then
        hypergeometric = [#VALUE!]
        Exit Function
     End If
     If ((i >= 0#) And ((ji <= 0#) Or (ki <= 0#)) Or (ip1 + ki <= 0#) Or (ip1 + ji <= 0#)) Then
        exact = True
        If (i >= 0#) Then
           prob = 1#
        Else
           prob = 0#
        End If
     ElseIf (ip1 > 0#) And (ip1 < 1#) Then
        exact = False
        prob = hypergeometricTerm(i, ji, ki, ip1) * (ip1 * (c4 + 1#)) / ((ki + ip1) * (ji + ip1))
     Else
        exact = ((i = 0#) And ((ji <= 0#) Or (ki <= 0#) Or (mkji + ki < 0#) Or (mkji + ji < 0#))) Or ((i > 0#) And (Min(ki, ji) = 0#) And (Max(mkji + ki, mkji + ji) = 0#))
        prob = hypergeometricTerm(i, ji, ki, mkji)
     End If
     hprob = prob
     hswap = swapped
     ha1 = 0#

     If (exact) Or (prob = 0#) Then
        If (swapped = comp) Then
           hypergeometric = prob
        Else
           hypergeometric = 1# - prob
        End If
        Exit Function
     End If

     a1 = 0#
     Dim sumAlways As Long, sumFactor As Long
     sumAlways = 0#
     sumFactor = 10#

     If i < mkji Then
        must_do_cf = i <> Int(i)
        maxSums = Int(i)
     Else
        must_do_cf = mkji <> Int(mkji)
        maxSums = Int(Max(mkji, 0#))
     End If
     If must_do_cf Then
        sumAlways = 0#
        sumFactor = 5#
     Else
        sumAlways = 20#
        sumFactor = 10#
     End If
     If (maxSums > sumAlways Or must_do_cf) Then
        numb = Int(sumFactor / c4 * Exp(Log((ki + i) * (ji + i) * (ip1 + ji) * (ip1 + ki)) / 3#))
        numb = Int(i - (ki + i) * (ji + i) / c4 + numb)
        If (numb < 0#) Then
           numb = 0#
        ElseIf numb > maxSums Then
           numb = maxSums
        End If
     Else
        numb = maxSums
     End If

     If (2# * numb <= maxSums Or must_do_cf) Then
        b1 = 1#
        c1 = 0#
        c2 = i - numb
        c3 = mkji - numb
        s = c3
        a2 = c2
        c3 = c3 - 1#
        b2 = GeneralabMinuscd(ki + numb + 1#, ji + numb + 1#, c2 - 1#, c3)
        bn = b2
        bnAdd = c3 + c4 + c2 - 2#
        While (b2 > 0# And (Abs(a2 * b1 - a1 * b2) > Abs(cfVSmall * b1 * a2)))
            c1 = c1 + 1#
            c2 = c2 - 1#
            an = (c1 * c2) * (c3 * c4)
            c3 = c3 - 1#
            c4 = c4 - 1#
            bn = bn + bnAdd
            bnAdd = bnAdd - 4#
            a1 = bn * a2 + an * a1
            b1 = bn * b2 + an * b1
            If (b1 > scalefactor) Then
              a1 = a1 * scalefactor2
              b1 = b1 * scalefactor2
              a2 = a2 * scalefactor2
              b2 = b2 * scalefactor2
            End If
            c1 = c1 + 1#
            c2 = c2 - 1#
            an = (c1 * c2) * (c3 * c4)
            c3 = c3 - 1#
            c4 = c4 - 1#
            bn = bn + bnAdd
            bnAdd = bnAdd - 4#
            a2 = bn * a1 + an * a2
            b2 = bn * b1 + an * b2
            If (b2 > scalefactor) Then
              a1 = a1 * scalefactor2
              b1 = b1 * scalefactor2
              a2 = a2 * scalefactor2
              b2 = b2 * scalefactor2
            End If
        Wend
        If b1 < 0# Or b2 < 0# Then
           hypergeometric = [#VALUE!]
           Exit Function
        Else
           a1 = a2 / b2 * s
        End If
     Else
        numb = maxSums
     End If

     c1 = i - numb + 1#
     c2 = mkji - numb + 1#
     c3 = ki + numb
     c4 = ji + numb
     For njj = 1 To numb
       a1 = (1# + a1) * ((c1 * c2) / (c3 * c4))
       c1 = c1 + 1#
       c2 = c2 + 1#
       c3 = c3 - 1#
       c4 = c4 - 1#
     Next njj

     ha1 = a1
     a1 = (1# + a1) * prob
     If (swapped = comp) Then
        hypergeometric = a1
     Else
        If a1 > 0.99 Then
           hypergeometric = [#VALUE!]
        Else
           hypergeometric = 1# - a1
        End If
     End If
End Function

Private Function compgfunc(ByVal x As Double, ByVal a As Double) As Double
'//Calculates a*x(1/(a+1) - x/2*(1/(a+2) - x/3*(1/(a+3) - ...)))
'//Mainly for calculating the complement of gamma(x,a) for small a and x <= 1.
'//a should be close to 0, x >= 0 & x <=1
  Dim term As Double, d As Double, sum As Double
  term = x
  d = 2#
  sum = term / (a + 1#)
  While (Abs(term) > Abs(sum * sumAcc))
      term = -term * x / d
      sum = sum + term / (a + d)
      d = d + 1#
  Wend
  compgfunc = a * sum
End Function

Private Function lngammaexpansion(ByVal a As Double) As Double
'//Calculates log(gamma(a+1)) accurately for for small a (a < 1.5).
'//Uses Abramowitz & Stegun's series 6.1.33
'//Mainly for calculating the complement of gamma(x,a) for small a and x <= 1.
'//
Dim lgam As Double
Dim i As Integer
Dim big As Boolean
Call initCoeffs
big = a > 0.5
If (big) Then
   a = a - 1#
End If
i = UBound(coeffs)
lgam = coeffs(i) * logcf(-a / 2#, i + 2#, 1#)
'More accurate with next line for larger values of a
'lgam = logcf(-a / 2#, i + 2#, 1#) * (2# ^ (-i - 2)) + logcf(-a / 3#, i + 2#, 1#) * (3# ^ (-i - 2))
For i = UBound(coeffs) - 1 To 0 Step -1
   lgam = (coeffs(i) - a * lgam)
Next i
lngammaexpansion = (a * lgam + OneMinusEulers_const) * a
If Not big Then
   lngammaexpansion = lngammaexpansion - log0(a)
End If
End Function

Private Function incgamma(ByVal x As Double, ByVal a As Double, ByVal comp As Boolean) As Double
'//Calculates gamma-cdf for small a (complementary gamma-cdf if comp).
   Dim r As Double
   r = a * Log(x) - lngammaexpansion(a)
   If (comp) Then
      r = -expm1(r)
      incgamma = r + compgfunc(x, a) * (1# - r)
   Else
      incgamma = Exp(r) * (1# - compgfunc(x, a))
   End If
End Function

Private Function invincgamma(ByVal a As Double, ByVal prob As Double, ByVal comp As Boolean) As Double
'//Calculates inverse of gamma for small a (inverse of complementary gamma if comp).
Dim ga As Double, x As Double, deriv As Double, z As Double, w As Double, dif As Double, pr As Double, lpr As Double, small As Double, smalllpr As Double
   If (prob > 0.5) Then
       invincgamma = invincgamma(a, 1# - prob, Not comp)
       Exit Function
   End If
   lpr = -Log(prob)
   small = 0.00000000000001
   smalllpr = small * lpr * prob
   If (comp) Then
      ga = -expm1(lngammaexpansion(a))
      x = -Log(prob * (1# - ga) / a)
      If (x < 0.5) Then
         pr = Exp(log0(-(ga + prob * (1# - ga))) / a)
         If (x < pr) Then
            x = pr
         End If
      End If
      dif = x
      pr = -1#
      While ((Abs(pr - prob) > smalllpr) And (Abs(dif) > small * Max(cSmall, x)))
         deriv = poissonTerm(a, x, x - a, 0#) * a             'value of derivative is actually deriv/x but it can overflow when x is denormal...
         If (x > 1#) Then
            pr = poisson1(-a, x, 0#)
         Else
            z = compgfunc(x, a)
            w = -expm1(a * Log(x))
            w = z + w * (1# - z)
            pr = (w - ga) / (1# - ga)
         End If
         dif = x * (pr / deriv) * logdif(pr, prob) '...so multiply by x in slightly different order
         x = x + dif
         If (x < 0#) Then
            invincgamma = 0#
            Exit Function
         End If
      Wend
   Else
      ga = Exp(lngammaexpansion(a))
      x = Log(prob * ga)
      If (x < -711# * a) Then
         invincgamma = 0#
         Exit Function
      End If
      x = Exp(x / a)
      z = 1# - compgfunc(x, a)
      deriv = poissonTerm(a, x, x - a, 0#) * a / x
      pr = prob * z
      dif = (pr / deriv) * logdif(pr, prob)
      x = x - dif
      While ((Abs(pr - prob) > smalllpr) And (Abs(dif) > small * Max(cSmall, x)))
         deriv = poissonTerm(a, x, x - a, 0#) * a / x
         If (x > 1#) Then
            pr = 1# - poisson1(-a, x, 0#)
         Else
            pr = (1# - compgfunc(x, a)) * Exp(a * Log(x)) / ga
         End If
         dif = (pr / deriv) * logdif(pr, prob)
         x = x - dif
      Wend
   End If
   invincgamma = x
End Function

Private Function gamma(ByVal n As Double, ByVal a As Double) As Double
'Assumes n > 0 & a >= 0.  Only called by (comp)gamma_nc with a = 0.
   If (a = 0#) Then
      gamma = 1#
   ElseIf ((a < 1#) And (n < 1#)) Then
      gamma = incgamma(n, a, False)
   ElseIf (a >= 1#) Then
      gamma = comppoisson(a - 1#, n, n - a + 1#)
   Else
      gamma = 1# - poisson1(-a, n, 0#)
   End If
End Function

Private Function compgamma(ByVal n As Double, ByVal a As Double) As Double
'Assumes n > 0 & a >= 0. Only called by (comp)gamma_nc with a = 0.
   If (a = 0#) Then
      compgamma = 0#
   ElseIf ((a < 1#) And (n < 1#)) Then
      compgamma = incgamma(n, a, True)
   ElseIf (a >= 1#) Then
      compgamma = cpoisson(a - 1#, n, n - a + 1#)
   Else
      compgamma = poisson1(-a, n, 0#)
   End If
End Function

Private Function invgamma(ByVal a As Double, ByVal prob As Double) As Double
'//Inverse of gamma(x,a)
   If (a >= 1#) Then
      invgamma = invcomppoisson(a - 1#, prob)
   Else
      invgamma = invincgamma(a, prob, False)
   End If
End Function

Private Function invcompgamma(ByVal a As Double, ByVal prob As Double) As Double
'//Inverse of compgamma(x,a)
   If (a >= 1#) Then
      invcompgamma = invpoisson(a - 1#, prob)
   Else
      invcompgamma = invincgamma(a, prob, True)
   End If
End Function

Private Function logfbit1dif(ByVal x As Double) As Double
'// Calculation of logfbit1(x)-logfbit1(1+x).
  'logfbit1dif = log0(1# / (x + 1#)) - (x + 1.5) / ((x + 1#) * (x + 2#))
  logfbit1dif = (logfbitdif(x) - 0.25 / ((x + 1#) * (x + 2#))) / (x + 1.5)
End Function

Private Function logfbit1(ByVal x As Double) As Double
'// Derivative of error part of Stirling's formula where log(x!) = log(sqrt(twopi))+(x+0.5)*log(x+1)-(x+1)+logfbit(x).
  Dim x1 As Double, x2 As Double
  If (x >= 10000000000#) Then
     logfbit1 = -lfbc1 * ((x + 1#) ^ -2)
  ElseIf (x >= 7#) Then
     Dim x3 As Double
     x1 = x + 1#
     x2 = 1# / (x1 * x1)
     x3 = (11# * lfbc6 - x2 * (13# * lfbc7 - x2 * (15# * lfbc8 - x2 * 17# * lfbc9)))
     x3 = (5# * lfbc3 - x2 * (7# * lfbc4 - x2 * (9# * lfbc5 - x2 * x3)))
     x3 = x2 * (3# * lfbc2 - x2 * x3)
    logfbit1 = -lfbc1 * (1# - x3) * x2
  ElseIf (x > -1#) Then
     x1 = x
     x2 = 0#
     While (x1 < 7#)
        x2 = x2 + logfbit1dif(x1)
        x1 = x1 + 1#
     Wend
     logfbit1 = x2 + logfbit1(x1)
  Else
     logfbit1 = -1E+308
  End If
End Function

Private Function logfbit2dif(ByVal x As Double) As Double
'// Calculation of logfbit2(x)-logfbit2(1+x).
  logfbit2dif = 0.5 * (((x + 1#) * (x + 2#)) ^ -2)
End Function

Private Function logfbit2(ByVal x As Double) As Double
'// Second derivative of error part of Stirling's formula where log(x!) = log(sqrt(twopi))+(x+0.5)*log(x+1)-(x+1)+logfbit(x).
  Dim x1 As Double, x2 As Double
  If (x >= 10000000000#) Then
     logfbit2 = 2# * lfbc1 * ((x + 1#) ^ -3)
  ElseIf (x >= 7#) Then
     Dim x3 As Double
     x1 = x + 1#
     x2 = 1# / (x1 * x1)
     x3 = x2 * (240# * lfbc8 - x2 * 306# * lfbc9)
     x3 = x2 * (132# * lfbc6 - x2 * (182# * lfbc7 - x3))
     x3 = x2 * (56# * lfbc4 - x2 * (90# * lfbc5 - x3))
     x3 = x2 * (12# * lfbc2 - x2 * (30# * lfbc3 - x3))
     logfbit2 = lfbc1 * (2# - x3) * x2 / x1
  ElseIf (x > -1#) Then
     x1 = x
     x2 = 0#
     While (x1 < 7#)
        x2 = x2 + logfbit2dif(x1)
        x1 = x1 + 1#
     Wend
     logfbit2 = x2 + logfbit2(x1)
  Else
     logfbit2 = -1E+308
  End If
End Function

Private Function logfbit3dif(ByVal x As Double) As Double
'// Calculation of logfbit3(x)-logfbit3(1+x).
  logfbit3dif = -(2# * x + 3#) * (((x + 1#) * (x + 2#)) ^ -3)
End Function

Private Function logfbit3(ByVal x As Double) As Double
'// Third derivative of error part of Stirling's formula where log(x!) = log(sqrt(twopi))+(x+0.5)*log(x+1)-(x+1)+logfbit(x).
  Dim x1 As Double, x2 As Double
  If (x >= 10000000000#) Then
     logfbit3 = -0.5 * ((x + 1#) ^ -4)
  ElseIf (x >= 7#) Then
     Dim x3 As Double
     x1 = x + 1#
     x2 = 1# / (x1 * x1)
     x3 = x2 * (4080# * lfbc8 - x2 * 5814# * lfbc9)
     x3 = x2 * (1716# * lfbc6 - x2 * (2730# * lfbc7 - x3))
     x3 = x2 * (504# * lfbc4 - x2 * (990# * lfbc5 - x3))
     x3 = x2 * (60# * lfbc2 - x2 * (210# * lfbc3 - x3))
     logfbit3 = -lfbc1 * (6# - x3) * x2 * x2
  ElseIf (x > -1#) Then
     x1 = x
     x2 = 0#
     While (x1 < 7#)
        x2 = x2 + logfbit3dif(x1)
        x1 = x1 + 1#
     Wend
     logfbit3 = x2 + logfbit3(x1)
  Else
     logfbit3 = -1E+308
  End If
End Function

Private Function logfbit4dif(ByVal x As Double) As Double
'// Calculation of logfbit4(x)-logfbit4(1+x).
  logfbit4dif = (10# * x * (x + 3#) + 23#) * (((x + 1#) * (x + 2#)) ^ -4)
End Function

Private Function logfbit4(ByVal x As Double) As Double
'// Fourth derivative of error part of Stirling's formula where log(x!) = log(sqrt(twopi))+(x+0.5)*log(x+1)-(x+1)+logfbit(x).
  Dim x1 As Double, x2 As Double
  If (x >= 10000000000#) Then
     logfbit4 = -0.5 * ((x + 1#) ^ -4)
  ElseIf (x >= 7#) Then
     Dim x3 As Double
     x1 = x + 1#
     x2 = 1# / (x1 * x1)
     x3 = x2 * (73440# * lfbc8 - x2 * 116280# * lfbc9)
     x3 = x2 * (24024# * lfbc6 - x2 * (43680# * lfbc7 - x3))
     x3 = x2 * (5040# * lfbc4 - x2 * (11880# * lfbc5 - x3))
     x3 = x2 * (360# * lfbc2 - x2 * (1680# * lfbc3 - x3))
     logfbit4 = lfbc1 * (24# - x3) * x2 * x2 / x1
  ElseIf (x > -1#) Then
     x1 = x
     x2 = 0#
     While (x1 < 7#)
        x2 = x2 + logfbit4dif(x1)
        x1 = x1 + 1#
     Wend
     logfbit4 = x2 + logfbit4(x1)
  Else
     logfbit4 = -1E+308
  End If
End Function

Private Function logfbit5dif(ByVal x As Double) As Double
'// Calculation of logfbit5(x)-logfbit5(1+x).
  logfbit5dif = -6# * (2# * x + 3#) * ((5# * x + 15#) * x + 12#) * (((x + 1#) * (x + 2#)) ^ -5)
End Function

Private Function logfbit5(ByVal x As Double) As Double
'// Fifth derivative of error part of Stirling's formula where log(x!) = log(sqrt(twopi))+(x+0.5)*log(x+1)-(x+1)+logfbit(x).
  Dim x1 As Double, x2 As Double
  If (x >= 10000000000#) Then
     logfbit5 = -10# * ((x + 1#) ^ -6)
  ElseIf (x >= 7#) Then
     Dim x3 As Double
     x1 = x + 1#
     x2 = 1# / (x1 * x1)
     x3 = x2 * (1395360# * lfbc8 - x2 * 2441880# * lfbc9)
     x3 = x2 * (360360# * lfbc6 - x2 * (742560# * lfbc7 - x3))
     x3 = x2 * (55440# * lfbc4 - x2 * (154440# * lfbc5 - x3))
     x3 = x2 * (2520# * lfbc2 - x2 * (15120# * lfbc3 - x3))
     logfbit5 = -lfbc1 * (120# - x3) * x2 * x2 * x2
  ElseIf (x > -1#) Then
     x1 = x
     x2 = 0#
     While (x1 < 7#)
        x2 = x2 + logfbit5dif(x1)
        x1 = x1 + 1#
     Wend
     logfbit5 = x2 + logfbit5(x1)
  Else
     logfbit5 = -1E+308
  End If
End Function

Private Function logfbit7dif(ByVal x As Double) As Double
'// Calculation of logfbit7(x)-logfbit7(1+x).
  logfbit7dif = -120# * (2# * x + 3#) * ((((14# * x + 84#) * x + 196#) * x + 210#) * x + 87#) * (((x + 1#) * (x + 2#)) ^ -7)
End Function

Private Function logfbit7(ByVal x As Double) As Double
'// Seventh derivative of error part of Stirling's formula where log(x!) = log(sqrt(twopi))+(x+0.5)*log(x+1)-(x+1)+logfbit(x).
  Dim x1 As Double, x2 As Double
  If (x >= 10000000000#) Then
     logfbit7 = -420# * ((x + 1#) ^ -8)
  ElseIf (x >= 7#) Then
     Dim x3 As Double
     x1 = x + 1#
     x2 = 1# / (x1 * x1)
     x3 = x2 * (586051200# * lfbc8 - x2 * 1235591280# * lfbc9)
     x3 = x2 * (98017920# * lfbc6 - x2 * (253955520# * lfbc7 - x3))
     x3 = x2 * (8648640# * lfbc4 - x2 * (32432400# * lfbc5 - x3))
     x3 = x2 * (181440# * lfbc2 - x2 * (1663200# * lfbc3 - x3))
     logfbit7 = -lfbc1 * (5040# - x3) * x2 * x2 * x2 * x2
  ElseIf (x > -1#) Then
     x1 = x
     x2 = 0#
     While (x1 < 7#)
        x2 = x2 + logfbit7dif(x1)
        x1 = x1 + 1#
     Wend
     logfbit7 = x2 + logfbit7(x1)
  Else
     logfbit7 = -1E+308
  End If
End Function

Private Function lfbaccdif(ByVal a As Double, ByVal b As Double) As Double
'// This is now always reasonably accurate, although it is not always required to be so when called from incbeta.
   If (a > 0.025 * (a + b + 1#)) Then
      lfbaccdif = logfbit(a + b) - logfbit(b)
   Else
      Dim a2 As Double, ab2 As Double
      a2 = a * a
      ab2 = a / 2# + b
      lfbaccdif = a * (logfbit1(ab2) + a2 / 24# * (logfbit3(ab2) + a2 / 80# * (logfbit5(ab2) + a2 / 168# * logfbit7(ab2))))
   End If
End Function

Private Function compbfunc(ByVal x As Double, ByVal a As Double, ByVal b As Double) As Double
'// Calculates a*(b-1)*x(1/(a+1) - (b-2)*x/2*(1/(a+2) - (b-3)*x/3*(1/(a+3) - ...)))
'// Mainly for calculating the complement of beta(x,a,b) for small a and b*x < 1.
'// a should be close to 0, x >= 0 & x <=1 & b*x < 1
  Dim term As Double, d As Double, sum As Double
  term = x
  d = 2#
  sum = term / (a + 1#)
  While (Abs(term) > Abs(sum * sumAcc))
      term = -term * (b - d) * x / d
      sum = sum + term / (a + d)
      d = d + 1#
  Wend
  compbfunc = a * (b - 1#) * sum
End Function

Private Function incbeta(ByVal x As Double, ByVal a As Double, ByVal b As Double, ByVal comp As Boolean) As Double
'// Calculates beta for small a (complementary beta if comp).
   Dim r As Double
   If (x > 0.5) Then
      incbeta = incbeta(1# - x, b, a, Not comp)
   Else
      r = (a + b + 0.5) * log1(a / (1# + b)) + a * ((a - 0.5) / (1# + b) + Log((1# + b) * x)) - lfbaccdif1(a, b) - lngammaexpansion(a)
      If (comp) Then
         r = -expm1(r)
         r = r + compbfunc(x, a, b) * (1# - r)
         r = r + (a / (a + b)) * (1# - r)
      Else
         r = Exp(r) * (1# - compbfunc(x, a, b)) * (b / (a + b))
      End If
      incbeta = r
   End If
End Function

Private Function beta(ByVal x As Double, ByVal a As Double, ByVal b As Double) As Double
'//Assumes x >= 0 & a >= 0 & b >= 0. Only called with a = 0 or b = 0 by (comp)beta_nc
   If (a = 0# And b = 0#) Then
      beta = [#VALUE!]
   ElseIf (a = 0#) Then
      beta = 1#
   ElseIf (b = 0#) Then
      beta = 0#
   ElseIf (x <= 0#) Then
      beta = 0#
   ElseIf (x >= 1#) Then
      beta = 1#
   ElseIf (a < 1# And b < 1#) Then
      beta = incbeta(x, a, b, False)
   ElseIf (a < 1# And (1# + b) * x <= 1#) Then
      beta = incbeta(x, a, b, False)
   ElseIf (b < 1# And a <= (1# + a) * x) Then
      beta = incbeta(1# - x, b, a, True)
   ElseIf (a < 1#) Then
      beta = compbinomial(-a, b, x, 1# - x, 0#)
   ElseIf (b < 1#) Then
      beta = binomial(-b, a, 1# - x, x, 0#)
   Else
      beta = compbinomial(a - 1#, b, x, 1# - x, (a + b - 1#) * x - a + 1#)
   End If
End Function

Private Function compbeta(ByVal x As Double, ByVal a As Double, ByVal b As Double) As Double
'//Assumes x >= 0 & a >= 0 & b >= 0. Only called with a = 0 or b = 0 by (comp)beta_nc
   If (a = 0# And b = 0#) Then
      compbeta = [#VALUE!]
   ElseIf (a = 0#) Then
      compbeta = 0#
   ElseIf (b = 0#) Then
      compbeta = 1#
   ElseIf (x <= 0#) Then
      compbeta = 1#
   ElseIf (x >= 1#) Then
      compbeta = 0#
   ElseIf (a < 1# And b < 1#) Then
      compbeta = incbeta(x, a, b, True)
   ElseIf (a < 1# And (1# + b) * x <= 1#) Then
      compbeta = incbeta(x, a, b, True)
   ElseIf (b < 1# And a <= (1# + a) * x) Then
      compbeta = incbeta(1# - x, b, a, False)
   ElseIf (a < 1#) Then
      compbeta = binomial(-a, b, x, 1# - x, 0#)
   ElseIf (b < 1#) Then
      compbeta = compbinomial(-b, a, 1# - x, x, 0#)
   Else
      compbeta = binomial(a - 1#, b, x, 1# - x, (a + b - 1#) * x - a + 1#)
   End If
End Function

Private Function invincbeta(ByVal a As Double, ByVal b As Double, ByVal prob As Double, ByVal comp As Boolean, ByRef oneMinusP As Double) As Double
'// Calculates inverse of beta for small a (inverse of complementary beta if comp).
Dim r As Double, rb As Double, x As Double, OneOverDeriv As Double, dif As Double, pr As Double, mnab As Double, aplusbOvermxab As Double, lpr As Double, small As Double, smalllpr As Double
   If (Not comp And prob > b / (a + b)) Then
       invincbeta = invincbeta(a, b, 1# - prob, Not comp, oneMinusP)
       Exit Function
   ElseIf (comp And prob > a / (a + b) And prob > 0.1) Then
       invincbeta = invincbeta(a, b, 1# - prob, Not comp, oneMinusP)
       Exit Function
   End If
   lpr = Max(-Log(prob), 1#)
   small = 0.00000000000001
   smalllpr = small * lpr * prob
   If a >= b Then
      mnab = b
      aplusbOvermxab = (a + b) / a
   Else
      mnab = a
      aplusbOvermxab = (a + b) / b
   End If
   If (comp) Then
      r = (a + b + 0.5) * log1(a / (1# + b)) + a * (a - 0.5) / (1# + b) - lfbaccdif1(a, b) - lngammaexpansion(a)
      r = -expm1(r)
      r = r + (a / (a + b)) * (1# - r)
      If (b < 1#) Then
         rb = (a + b + 0.5) * log1(b / (1# + a)) + b * (b - 0.5) / (1# + a) - lfbaccdif1(b, a) - lngammaexpansion(b)
         rb = Exp(rb) * (a / (a + b))
         oneMinusP = Log(prob / rb) / b
         If (oneMinusP < 0#) Then
             oneMinusP = Exp(oneMinusP) / (1# + a)
         Else
             oneMinusP = 0.5
         End If
         If (oneMinusP = 0#) Then
            invincbeta = 1#
            Exit Function
         ElseIf (oneMinusP > 0.5) Then
            oneMinusP = 0.5
         End If
         x = 1# - oneMinusP
         pr = rb * (1# - compbfunc(oneMinusP, b, a)) * Exp(b * Log((1# + a) * oneMinusP))
         OneOverDeriv = (aplusbOvermxab * x * oneMinusP) / (binomialTerm(a, b, x, oneMinusP, (a + b) * x - a, 0#) * mnab)
         dif = OneOverDeriv * pr * logdif(pr, prob)
         oneMinusP = oneMinusP - dif
         x = 1# - oneMinusP
         If (oneMinusP <= 0#) Then
            oneMinusP = 0#
            invincbeta = 1#
            Exit Function
         ElseIf (x < 0.25) Then
            x = Exp(log0((r - prob) / (1# - r)) / a) / (b + 1#)
            oneMinusP = 1# - x
            If (x = 0#) Then
               invincbeta = 0#
               Exit Function
            End If
            pr = compbfunc(x, a, b) * (1# - prob)
            OneOverDeriv = (aplusbOvermxab * x * oneMinusP) / (binomialTerm(a, b, x, oneMinusP, (a + b) * x - a, 0#) * mnab)
            dif = OneOverDeriv * (prob + pr) * log0(pr / prob)
            x = x + dif
            If (x <= 0#) Then
               oneMinusP = 1#
               invincbeta = 0#
               Exit Function
            End If
            oneMinusP = 1# - x
         End If
      Else
         pr = Exp(log0((r - prob) / (1# - r)) / a) / (b + 1#)
         x = Log(b * prob / (a * (1# - r) * b * Exp(a * Log(1# + b)))) / b
         If (Abs(x) < 0.5) Then
            x = -expm1(x)
            oneMinusP = 1# - x
         Else
            oneMinusP = Exp(x)
            x = 1# - oneMinusP
            If (oneMinusP = 0#) Then
               invincbeta = x
               Exit Function
            End If
         End If
         If pr > x And pr < 1# Then
            x = pr
            oneMinusP = 1# - x
         End If
      End If
      dif = Min(x, oneMinusP)
      pr = -1#
      While ((Abs(pr - prob) > smalllpr) And (Abs(dif) > small * Max(cSmall, Min(x, oneMinusP))))
         If (b < 1# And x > 0.5) Then
            pr = rb * (1# - compbfunc(oneMinusP, b, a)) * Exp(b * Log((1# + a) * oneMinusP))
         ElseIf ((1# + b) * x > 1#) Then
            pr = binomial(-a, b, x, oneMinusP, 0#)
         Else
            pr = r + compbfunc(x, a, b) * (1# - r)
            pr = pr - expm1(a * Log((1# + b) * x)) * (1# - pr)
         End If
         OneOverDeriv = (aplusbOvermxab * x * oneMinusP) / (binomialTerm(a, b, x, oneMinusP, (a + b) * x - a, 0#) * mnab)
         dif = OneOverDeriv * pr * logdif(pr, prob)
         If (x > 0.5) Then
            oneMinusP = oneMinusP - dif
            x = 1# - oneMinusP
            If (oneMinusP <= 0#) Then
               oneMinusP = 0#
               invincbeta = 1#
               Exit Function
            End If
         Else
            x = x + dif
            oneMinusP = 1# - x
            If (x <= 0#) Then
               oneMinusP = 1#
               invincbeta = 0#
               Exit Function
            End If
         End If
      Wend
   Else
      r = (a + b + 0.5) * log1(a / (1# + b)) + a * (a - 0.5) / (1# + b) - lfbaccdif1(a, b) - lngammaexpansion(a)
      r = Exp(r) * (b / (a + b))
      x = logdif(prob, r)
      If (x < -711# * a) Then
         x = 0#
      Else
         x = Exp(x / a) / (1# + b)
      End If
      If (x = 0#) Then
         oneMinusP = 1#
         invincbeta = x
         Exit Function
      ElseIf (x >= 0.5) Then
         x = 0.5
      End If
      oneMinusP = 1# - x
      pr = r * (1# - compbfunc(x, a, b)) * Exp(a * Log((1# + b) * x))
      OneOverDeriv = (aplusbOvermxab * x * oneMinusP) / (binomialTerm(a, b, x, oneMinusP, (a + b) * x - a, 0#) * mnab)
      dif = OneOverDeriv * pr * logdif(pr, prob)
      x = x - dif
      oneMinusP = oneMinusP + dif
      While ((Abs(pr - prob) > smalllpr) And (Abs(dif) > small * Max(cSmall, Min(x, oneMinusP))))
         If ((1# + b) * x > 1#) Then
            pr = compbinomial(-a, b, x, oneMinusP, 0#)
         ElseIf (x > 0.5) Then
            pr = incbeta(oneMinusP, b, a, Not comp)
         Else
            pr = r * (1# - compbfunc(x, a, b)) * Exp(a * Log((1# + b) * x))
         End If
         OneOverDeriv = (aplusbOvermxab * x * oneMinusP) / (binomialTerm(a, b, x, oneMinusP, (a + b) * x - a, 0#) * mnab)
         dif = OneOverDeriv * pr * logdif(pr, prob)
         If x < 0.5 Then
            x = x - dif
            oneMinusP = 1# - x
         Else
            oneMinusP = oneMinusP + dif
            x = 1# - oneMinusP
         End If
      Wend
   End If
   invincbeta = x
End Function

Private Function invbeta(ByVal a As Double, ByVal b As Double, ByVal prob As Double, ByRef oneMinusP As Double) As Double
   Dim swap As Double
   If (prob = 0#) Then
      oneMinusP = 1#
      invbeta = 0#
   ElseIf (prob = 1#) Then
      oneMinusP = 0#
      invbeta = 1#
   ElseIf (a = b And prob = 0.5) Then
      oneMinusP = 0.5
      invbeta = 0.5
   ElseIf (a < b And b < 1#) Then
      invbeta = invincbeta(a, b, prob, False, oneMinusP)
   ElseIf (b < a And a < 1#) Then
      swap = invincbeta(b, a, prob, True, oneMinusP)
      invbeta = oneMinusP
      oneMinusP = swap
   ElseIf (a < 1#) Then
      invbeta = invincbeta(a, b, prob, False, oneMinusP)
   ElseIf (b < 1#) Then
      swap = invincbeta(b, a, prob, True, oneMinusP)
      invbeta = oneMinusP
      oneMinusP = swap
   Else
      invbeta = invcompbinom(a - 1#, b, prob, oneMinusP)
   End If
End Function

Private Function invcompbeta(ByVal a As Double, ByVal b As Double, ByVal prob As Double, ByRef oneMinusP As Double) As Double
   Dim swap As Double
   If (prob = 0#) Then
      oneMinusP = 0#
      invcompbeta = 1#
   ElseIf (prob = 1#) Then
      oneMinusP = 1#
      invcompbeta = 0#
   ElseIf (a = b And prob = 0.5) Then
      oneMinusP = 0.5
      invcompbeta = 0.5
   ElseIf (a < b And b < 1#) Then
      invcompbeta = invincbeta(a, b, prob, True, oneMinusP)
   ElseIf (b < a And a < 1#) Then
      swap = invincbeta(b, a, prob, False, oneMinusP)
      invcompbeta = oneMinusP
      oneMinusP = swap
   ElseIf (a < 1#) Then
      invcompbeta = invincbeta(a, b, prob, True, oneMinusP)
   ElseIf (b < 1#) Then
      swap = invincbeta(b, a, prob, False, oneMinusP)
      invcompbeta = oneMinusP
      oneMinusP = swap
   Else
      invcompbeta = invbinom(a - 1#, b, prob, oneMinusP)
   End If
End Function

Private Function critpoiss(ByVal mean As Double, ByVal cprob As Double) As Double
'//i such that Pr(poisson(mean,i)) >= cprob and  Pr(poisson(mean,i-1)) < cprob
   If (cprob > 0.5) Then
      critpoiss = critcomppoiss(mean, 1# - cprob)
      Exit Function
   End If
   Dim pr As Double, tpr As Double, dfm As Double
   Dim i As Double
   dfm = invcnormal(cprob) * Sqr(mean)
   i = Int(mean + dfm + 0.5)
   While (True)
      i = Int(i)
      If (i < 0#) Then
         i = 0#
      End If
      If (i >= max_crit) Then
         critpoiss = i
         Exit Function
      End If
      dfm = mean - i
      pr = cpoisson(i, mean, dfm)
      tpr = 0#
      If (pr >= cprob) Then
         If (i = 0#) Then
            critpoiss = i
            Exit Function
         End If
         tpr = poissonTerm(i, mean, dfm, 0#)
         pr = pr - tpr
         If (pr < cprob) Then
            critpoiss = i
            Exit Function
         End If

         i = i - 1#
         Dim temp As Double, temp2 As Double
         temp = (pr - cprob) / tpr
         If (temp > 10) Then
            temp = Int(temp + 0.5)
            i = i - temp
            temp2 = poissonTerm(i, mean, mean - i, 0#)
            i = i - temp * (tpr - temp2) / (2 * temp2)
         Else
            tpr = tpr * (i + 1#) / mean
            pr = pr - tpr
            If (pr < cprob) Then
               critpoiss = i
               Exit Function
            End If
            i = i - 1#
            If (i = 0#) Then
               critpoiss = i
               Exit Function
            End If
            temp2 = (pr - cprob) / tpr
            If (temp2 < temp - 0.9) Then
               While (pr >= cprob)
                  tpr = tpr * (i + 1#) / mean
                  pr = pr - tpr
                  i = i - 1#
               Wend
               critpoiss = i + 1#
               Exit Function
            Else
               temp = Int(Log(cprob / pr) / Log((i + 1#) / mean) + 0.5)
               i = i - temp
               If (i < 0#) Then
                  i = 0#
               End If
               temp2 = poissonTerm(i, mean, mean - i, 0#)
               If (temp2 > nearly_zero) Then
                  temp = Log((cprob / pr) * (tpr / temp2)) / Log((i + 1#) / mean)
                  i = i - temp
               End If
            End If
         End If
      Else
         While ((tpr < cSmall) And (pr < cprob))
            i = i + 1#
            dfm = dfm - 1#
            tpr = poissonTerm(i, mean, dfm, 0#)
            pr = pr + tpr
         Wend
         While (pr < cprob)
            i = i + 1#
            tpr = tpr * mean / i
            pr = pr + tpr
         Wend
         critpoiss = i
         Exit Function
      End If
   Wend
End Function

Private Function critcomppoiss(ByVal mean As Double, ByVal cprob As Double) As Double
'//i such that 1-Pr(poisson(mean,i)) > cprob and  1-Pr(poisson(mean,i-1)) <= cprob
   If (cprob > 0.5) Then
      critcomppoiss = critpoiss(mean, 1# - cprob)
      Exit Function
   End If
   Dim pr As Double, tpr As Double, dfm As Double
   Dim i As Double
   dfm = invcnormal(cprob) * Sqr(mean)
   i = Int(mean - dfm + 0.5)
   While (True)
      i = Int(i)
      If (i >= max_crit) Then
         critcomppoiss = i
         Exit Function
      End If
      dfm = mean - i
      pr = comppoisson(i, mean, dfm)
      tpr = 0#
      If (pr > cprob) Then
         i = i + 1#
         dfm = dfm - 1#
         tpr = poissonTerm(i, mean, dfm, 0#)
         If (pr < (1.00001) * tpr) Then
            While (tpr > cprob)
               i = i + 1#
               tpr = tpr * mean / i
            Wend
         Else
            pr = pr - tpr
            If (pr <= cprob) Then
               critcomppoiss = i
               Exit Function
            End If
            Dim temp As Double, temp2 As Double
            temp = (pr - cprob) / tpr
            If (temp > 10) Then
               temp = Int(temp + 0.5)
               i = i + temp
               temp2 = poissonTerm(i, mean, mean - i, 0#)
               i = i + temp * (tpr - temp2) / (2# * temp2)
            ElseIf (pr / tpr > 0.00001) Then
               i = i + 1#
               tpr = tpr * mean / i
               pr = pr - tpr
               If (pr <= cprob) Then
                  critcomppoiss = i
                  Exit Function
               End If
               temp2 = (pr - cprob) / tpr
               If (temp2 < temp - 0.9) Then
                  While (pr > cprob)
                     i = i + 1#
                     tpr = tpr * mean / i
                     pr = pr - tpr
                  Wend
                  critcomppoiss = i
                  Exit Function
               Else
                  temp = Log(cprob / pr) / Log(mean / i)
                  temp = Int((Log(cprob / pr) - temp * Log(i / (temp + i))) / Log(mean / i) + 0.5)
                  i = i + temp
                  temp2 = poissonTerm(i, mean, mean - i, 0#)
                  If (temp2 > nearly_zero) Then
                     temp = Log((cprob / pr) * (tpr / temp2)) / Log(mean / i)
                     i = i + temp
                  End If
               End If
            End If
         End If
      Else
         While ((tpr < cSmall) And (pr <= cprob))
            tpr = poissonTerm(i, mean, dfm, 0#)
            pr = pr + tpr
            i = i - 1#
            dfm = dfm + 1#
         Wend
         While (pr <= cprob)
            tpr = tpr * (i + 1#) / mean
            pr = pr + tpr
            i = i - 1#
         Wend
         critcomppoiss = i + 1#
         Exit Function
      End If
   Wend
End Function

Private Function critbinomial(ByVal n As Double, ByVal eprob As Double, ByVal cprob As Double) As Double
'//i such that Pr(binomial(n,eprob,i)) >= cprob and  Pr(binomial(n,eprob,i-1)) < cprob
   If (cprob > 0.5) Then
      critbinomial = critcompbinomial(n, eprob, 1# - cprob)
      Exit Function
   End If
   Dim pr As Double, tpr As Double, dfm As Double
   Dim i As Double
   dfm = invcnormal(cprob) * Sqr(n * eprob * (1# - eprob))
   i = n * eprob + dfm
   While (True)
      i = Int(i)
      If (i < 0#) Then
         i = 0#
      ElseIf (i > n) Then
         i = n
      End If
      If (i >= max_crit) Then
         critbinomial = i
         Exit Function
      End If
      dfm = n * eprob - i
      pr = binomial(i, n - i, eprob, 1# - eprob, dfm)
      tpr = 0#
      If (pr >= cprob) Then
         If (i = 0#) Then
            critbinomial = i
            Exit Function
         End If
         tpr = binomialTerm(i, n - i, eprob, 1# - eprob, dfm, 0#)
         If (pr < (1.00001) * tpr) Then
            tpr = tpr * ((i + 1#) * (1# - eprob)) / ((n - i) * eprob)
            i = i - 1#
            While (tpr >= cprob)
               tpr = tpr * ((i + 1#) * (1# - eprob)) / ((n - i) * eprob)
               i = i - 1
            Wend
         Else
            pr = pr - tpr
            If (pr < cprob) Then
               critbinomial = i
               Exit Function
            End If
            i = i - 1#
            If (i = 0#) Then
               critbinomial = i
               Exit Function
            End If
            Dim temp As Double, temp2 As Double
            temp = (pr - cprob) / tpr
            If (temp > 10#) Then
               temp = Int(temp + 0.5)
               i = i - temp
               temp2 = binomialTerm(i, n - i, eprob, 1# - eprob, n * eprob - i, 0#)
               i = i - temp * (tpr - temp2) / (2# * temp2)
            Else
               tpr = tpr * ((i + 1#) * (1# - eprob)) / ((n - i) * eprob)
               pr = pr - tpr
               If (pr < cprob) Then
                  critbinomial = i
                  Exit Function
               End If
               i = i - 1#
               temp2 = (pr - cprob) / tpr
               If (temp2 < temp - 0.9) Then
                  While (pr >= cprob)
                     tpr = tpr * ((i + 1#) * (1# - eprob)) / ((n - i) * eprob)
                     pr = pr - tpr
                     i = i - 1#
                  Wend
                  critbinomial = i + 1#
                  Exit Function
               Else
                  temp = Int(Log(cprob / pr) / Log(((i + 1#) * (1# - eprob)) / ((n - i) * eprob)) + 0.5)
                  i = i - temp
                  If (i < 0#) Then
                     i = 0#
                  End If
                  temp2 = binomialTerm(i, n - i, eprob, 1# - eprob, n * eprob - i, 0#)
                  If (temp2 > nearly_zero) Then
                     temp = Log((cprob / pr) * (tpr / temp2)) / Log(((i + 1#) * (1# - eprob)) / ((n - i) * eprob))
                     i = i - temp
                  End If
               End If
            End If
         End If
      Else
         While ((tpr < cSmall) And (pr < cprob))
            i = i + 1#
            dfm = dfm - 1#
            tpr = binomialTerm(i, n - i, eprob, 1# - eprob, dfm, 0#)
            pr = pr + tpr
         Wend
         While (pr < cprob)
            i = i + 1#
            tpr = tpr * ((n - i + 1#) * eprob) / (i * (1# - eprob))
            pr = pr + tpr
         Wend
         critbinomial = i
         Exit Function
      End If
   Wend
End Function

Private Function critcompbinomial(ByVal n As Double, ByVal eprob As Double, ByVal cprob As Double) As Double
'//i such that 1-Pr(binomial(n,eprob,i)) > cprob and  1-Pr(binomial(n,eprob,i-1)) <= cprob
   If (cprob > 0.5) Then
      critcompbinomial = critbinomial(n, eprob, 1# - cprob)
      Exit Function
   End If
   Dim pr As Double, tpr As Double, dfm As Double
   Dim i As Double
   dfm = invcnormal(cprob) * Sqr(n * eprob * (1# - eprob))
   i = n * eprob - dfm
   While (True)
      i = Int(i)
      If (i < 0#) Then
         i = 0#
      ElseIf (i > n) Then
         i = n
      End If
      If (i >= max_crit) Then
         critcompbinomial = i
         Exit Function
      End If
      dfm = n * eprob - i
      pr = compbinomial(i, n - i, eprob, 1# - eprob, dfm)
      tpr = 0#
      If (pr > cprob) Then
         i = i + 1#
         dfm = dfm - 1#
         tpr = binomialTerm(i, n - i, eprob, 1# - eprob, dfm, 0#)
         If (pr < (1.00001) * tpr) Then
            While (tpr > cprob)
               i = i + 1#
               tpr = tpr * ((n - i + 1#) * eprob) / (i * (1# - eprob))
            Wend
         Else
            pr = pr - tpr
            If (pr <= cprob) Then
               critcompbinomial = i
               Exit Function
            End If
            Dim temp As Double, temp2 As Double
            temp = (pr - cprob) / tpr
            If (temp > 10#) Then
               temp = Int(temp + 0.5)
               i = i + temp
               temp2 = binomialTerm(i, n - i, eprob, 1# - eprob, n * eprob - i, 0#)
               i = i + temp * (tpr - temp2) / (2# * temp2)
            Else
               i = i + 1#
               tpr = tpr * ((n - i + 1#) * eprob) / (i * (1# - eprob))
               pr = pr - tpr
               If (pr <= cprob) Then
                  critcompbinomial = i
                  Exit Function
               End If
               temp2 = (pr - cprob) / tpr
               If (temp2 < temp - 0.9) Then
                  While (pr > cprob)
                     i = i + 1#
                     tpr = tpr * ((n - i + 1#) * eprob) / (i * (1# - eprob))
                     pr = pr - tpr
                  Wend
                  critcompbinomial = i
                  Exit Function
               Else
                  temp = Int(Log(cprob / pr) / Log(((n - i + 1#) * eprob) / (i * (1# - eprob))) + 0.5)
                  i = i + temp
                  If (i > n) Then
                     i = n
                  End If
                  temp2 = binomialTerm(i, n - i, eprob, 1# - eprob, n * eprob - i, 0#)
                  If (temp2 > nearly_zero) Then
                     temp = Log((cprob / pr) * (tpr / temp2)) / Log(((n - i + 1#) * eprob) / (i * (1# - eprob)))
                     i = i + temp
                  End If
               End If
            End If
         End If
      Else
         While ((tpr < cSmall) And (pr <= cprob))
            tpr = binomialTerm(i, n - i, eprob, 1# - eprob, dfm, 0#)
            pr = pr + tpr
            i = i - 1#
            dfm = dfm + 1#
         Wend
         While (pr <= cprob)
            tpr = tpr * ((i + 1#) * (1# - eprob)) / ((n - i) * eprob)
            pr = pr + tpr
            i = i - 1#
         Wend
         critcompbinomial = i + 1#
         Exit Function
      End If
   Wend
End Function

Private Function crithyperg(ByVal j As Double, ByVal k As Double, ByVal m As Double, ByVal cprob As Double) As Double
'//i such that Pr(hypergeometric(i,j,k,m)) >= cprob and  Pr(hypergeometric(i-1,j,k,m)) < cprob
   Dim ha1 As Double, hprob As Double, hswap As Boolean
   If (cprob > 0.5) Then
      crithyperg = critcomphyperg(j, k, m, 1# - cprob)
      Exit Function
   End If
   Dim pr As Double, tpr As Double
   Dim i As Double
   i = j * k / m + invcnormal(cprob) * Sqr(j * k * (m - j) * (m - k) / (m * m * (m - 1#)))
   Dim mx As Double, mn  As Double
   mx = Min(j, k)
   mn = Max(0, j + k - m)
   While (True)
      If (i < mn) Then
         i = mn
      ElseIf (i > mx) Then
         i = mx
      End If
      i = Int(i + 0.5)
      If (i >= max_crit) Then
         crithyperg = i
         Exit Function
      End If
      pr = hypergeometric(i, j - i, k - i, m - k - j + i, False, ha1, hprob, hswap)
      tpr = 0#
      If (pr >= cprob) Then
         If (i = mn) Then
            crithyperg = mn
            Exit Function
         End If
         tpr = hypergeometricTerm(i, j - i, k - i, m - k - j + i)
         If (pr < (1.00001) * tpr) Then
            tpr = tpr * ((i + 1#) * (m - j - k + i + 1#)) / ((k - i) * (j - i))
            i = i - 1#
            While (tpr > cprob)
               tpr = tpr * ((i + 1#) * (m - j - k + i + 1#)) / ((k - i) * (j - i))
               i = i - 1#
            Wend
         Else
            pr = pr - tpr
            If (pr < cprob) Then
               crithyperg = i
               Exit Function
            End If
            i = i - 1#
            If (i = mn) Then
               crithyperg = mn
               Exit Function
            End If
            Dim temp As Double, temp2 As Double
            temp = (pr - cprob) / tpr
            If (temp > 10) Then
               temp = Int(temp + 0.5)
               i = i - temp
               temp2 = hypergeometricTerm(i, j - i, k - i, m - k - j + i)
               i = i - temp * (tpr - temp2) / (2# * temp2)
            Else
               tpr = tpr * ((i + 1#) * (m - j - k + i + 1#)) / ((k - i) * (j - i))
               pr = pr - tpr
               If (pr < cprob) Then
                  crithyperg = i
                  Exit Function
               End If
               i = i - 1#
               temp2 = (pr - cprob) / tpr
               If (temp2 < temp - 0.9) Then
                  While (pr >= cprob)
                     tpr = tpr * ((i + 1#) * (m - j - k + i + 1#)) / ((k - i) * (j - i))
                     pr = pr - tpr
                     i = i - 1#
                  Wend
                  crithyperg = i + 1#
                  Exit Function
               Else
                  temp = Int(Log(cprob / pr) / Log(((i + 1#) * (m - j - k + i + 1#)) / ((k - i) * (j - i))) + 0.5)
                  i = i - temp
                  temp2 = hypergeometricTerm(i, j - i, k - i, m - k - j + i)
                  If (temp2 > nearly_zero) Then
                     temp = Log((cprob / pr) * (tpr / temp2)) / Log(((i + 1#) * (m - j - k + i + 1#)) / ((k - i) * (j - i)))
                     i = i - temp
                  End If
               End If
            End If
         End If
      Else
         While ((tpr < cSmall) And (pr < cprob))
            i = i + 1#
            tpr = hypergeometricTerm(i, j - i, k - i, m - k - j + i)
            pr = pr + tpr
         Wend
         While (pr < cprob)
            i = i + 1#
            tpr = tpr * ((k - i + 1#) * (j - i + 1#)) / (i * (m - j - k + i))
            pr = pr + tpr
         Wend
         crithyperg = i
         Exit Function
      End If
   Wend
End Function

Private Function critcomphyperg(ByVal j As Double, ByVal k As Double, ByVal m As Double, ByVal cprob As Double) As Double
'//i such that 1-Pr(hypergeometric(i,j,k,m)) > cprob and  1-Pr(hypergeometric(i-1,j,k,m)) <= cprob
   Dim ha1 As Double, hprob As Double, hswap As Boolean
   If (cprob > 0.5) Then
      critcomphyperg = crithyperg(j, k, m, 1# - cprob)
      Exit Function
   End If
   Dim pr As Double, tpr As Double
   Dim i As Double
   i = j * k / m - invcnormal(cprob) * Sqr(j * k * (m - j) * (m - k) / (m * m * (m - 1#)))
   Dim mx As Double, mn  As Double
   mx = Min(j, k)
   mn = Max(0, j + k - m)
   While (True)
      If (i < mn) Then
         i = mn
      ElseIf (i > mx) Then
         i = mx
      End If
      i = Int(i + 0.5)
      If (i >= max_crit) Then
         critcomphyperg = i
         Exit Function
      End If
      pr = hypergeometric(i, j - i, k - i, m - k - j + i, True, ha1, hprob, hswap)
      tpr = 0#
      If (pr > cprob) Then
         i = i + 1#
         tpr = hypergeometricTerm(i, j - i, k - i, m - k - j + i)
         If (pr < (1# + 0.00001) * tpr) Then
            While (tpr > cprob)
               i = i + 1
               tpr = tpr * ((k - i + 1#) * (j - i + 1#)) / (i * (m - j - k + i))
            Wend
         Else
            pr = pr - tpr
            If (pr <= cprob) Then
               critcomphyperg = i
               Exit Function
            End If
            Dim temp As Double, temp2 As Double
            temp = (pr - cprob) / tpr
            If (temp > 10) Then
               temp = Int(temp + 0.5)
               i = i + temp
               temp2 = hypergeometricTerm(i, j - i, k - i, m - k - j + i)
               i = i + temp * (tpr - temp2) / (2# * temp2)
            Else
               i = i + 1#
               tpr = tpr * ((k - i + 1#) * (j - i + 1#)) / (i * (m - j - k + i))
               pr = pr - tpr
               If (pr <= cprob) Then
                  critcomphyperg = i
                  Exit Function
               End If
               temp2 = (pr - cprob) / tpr
               If (temp2 < temp - 0.9) Then
                  While (pr > cprob)
                     i = i + 1#
                     tpr = tpr * ((k - i + 1#) * (j - i + 1#)) / (i * (m - j - k + i))
                     pr = pr - tpr
                  Wend
                  critcomphyperg = i
                  Exit Function
               Else
                  temp = Int(Log(cprob / pr) / Log(((k - i + 1#) * (j - i + 1#)) / (i * (m - j - k + i))) + 0.5)
                  i = i + temp
                  temp2 = hypergeometricTerm(i, j - i, k, m - k)
                  If (temp2 > nearly_zero) Then
                     temp = Log((cprob / pr) * (tpr / temp2)) / Log(((k - i + 1#) * (j - i + 1#)) / (i * (m - j - k + i)))
                     i = i + temp
                  End If
               End If
            End If
         End If
      Else
         While ((tpr < cSmall) And (pr <= cprob))
            tpr = hypergeometricTerm(i, j - i, k - i, m - k - j + i)
            pr = pr + tpr
            i = i - 1#
         Wend
         While (pr <= cprob)
            tpr = tpr * ((i + 1#) * (m - j - k + i + 1#)) / ((k - i) * (j - i))
            pr = pr + tpr
            i = i - 1#
         Wend
         critcomphyperg = i + 1#
         Exit Function
      End If
   Wend
End Function

Private Function critnegbinom(ByVal n As Double, ByVal eprob As Double, ByVal fprob As Double, ByVal cprob As Double) As Double
'//i such that Pr(negbinomial(n,eprob,i)) >= cprob and  Pr(negbinomial(n,eprob,i-1)) < cprob
   If (cprob > 0.5) Then
      critnegbinom = critcompnegbinom(n, eprob, fprob, 1# - cprob)
      Exit Function
   End If
   Dim pr As Double, tpr As Double, dfm As Double
   Dim i As Double
   i = invgamma(n * fprob, cprob) / eprob
   While (True)
      If (i < 0#) Then
         i = 0#
      End If
      i = Int(i)
      If (i >= max_crit) Then
         critnegbinom = i
         Exit Function
      End If
      If eprob <= fprob Then
         pr = beta(eprob, n, i + 1#)
      Else
         pr = compbeta(fprob, i + 1#, n)
      End If
      tpr = 0#
      If (pr >= cprob) Then
         If (i = 0#) Then
            critnegbinom = i
            Exit Function
         End If
         If eprob <= fprob Then
            dfm = n - (n + i) * eprob
         Else
            dfm = (n + i) * fprob - i
         End If
         tpr = n / (n + i) * binomialTerm(i, n, fprob, eprob, dfm, 0#)
         If (pr < (1.00001) * tpr) Then
            tpr = tpr * (i + 1#) / ((n + i) * fprob)
            i = i - 1#
            While (tpr > cprob)
               tpr = tpr * (i + 1#) / ((n + i) * fprob)
               i = i - 1#
            Wend
         Else
            pr = pr - tpr
            If (pr < cprob) Then
               critnegbinom = i
               Exit Function
            End If
            i = i - 1#
            If (i = 0#) Then
               critnegbinom = i
               Exit Function
            End If
            Dim temp As Double, temp2 As Double
            temp = (pr - cprob) / tpr
            If (temp > 10#) Then
               temp = Int(temp + 0.5)
               i = i - temp
               If eprob <= fprob Then
                  dfm = n - (n + i) * eprob
               Else
                  dfm = (n + i) * fprob - i
               End If
               temp2 = n / (n + i) * binomialTerm(i, n, fprob, eprob, dfm, 0#)
               i = i - temp * (tpr - temp2) / (2# * temp2)
            Else
               tpr = tpr * (i + 1#) / ((n + i) * fprob)
               pr = pr - tpr
               If (pr < cprob) Then
                  critnegbinom = i
                  Exit Function
               End If
               i = i - 1#
               temp2 = (pr - cprob) / tpr
               If (temp2 < temp - 0.9) Then
                  While (pr >= cprob)
                     tpr = tpr * (i + 1#) / ((n + i) * fprob)
                     pr = pr - tpr
                     i = i - 1#
                  Wend
                  critnegbinom = i + 1#
                  Exit Function
               Else
                  temp = Int(Log(cprob / pr) / Log((i + 1#) / ((n + i) * fprob)) + 0.5)
                  i = i - temp
                  If eprob <= fprob Then
                     dfm = n - (n + i) * eprob
                  Else
                     dfm = (n + i) * fprob - i
                  End If
                  temp2 = n / (n + i) * binomialTerm(i, n, fprob, eprob, dfm, 0#)
                  If (temp2 > nearly_zero) Then
                     temp = Log((cprob / pr) * (tpr / temp2)) / Log((i + 1#) / ((n + i) * fprob))
                     i = i - temp
                  End If
               End If
            End If
         End If
      Else
         While ((tpr < cSmall) And (pr < cprob))
            i = i + 1#
            If eprob <= fprob Then
               dfm = n - (n + i) * eprob
            Else
               dfm = (n + i) * fprob - i
            End If
            tpr = n / (n + i) * binomialTerm(i, n, fprob, eprob, dfm, 0#)
            pr = pr + tpr
         Wend
         While (pr < cprob)
            i = i + 1#
            tpr = tpr * ((n + i - 1#) * fprob) / i
            pr = pr + tpr
         Wend
         critnegbinom = i
         Exit Function
      End If
   Wend
End Function

Private Function critcompnegbinom(ByVal n As Double, ByVal eprob As Double, ByVal fprob As Double, ByVal cprob As Double) As Double
'//i such that 1-Pr(negbinomial(n,eprob,i)) > cprob and  1-Pr(negbinomial(n,eprob,i-1)) <= cprob
   If (cprob > 0.5) Then
      critcompnegbinom = critnegbinom(n, eprob, fprob, 1# - cprob)
      Exit Function
   End If
   Dim pr As Double, tpr As Double, dfm As Double
   Dim i As Double
   i = invcompgamma(n * fprob, cprob) / eprob
   While (True)
      If (i < 0#) Then
         i = 0#
      End If
      i = Int(i)
      If (i >= max_crit) Then
         critcompnegbinom = i
         Exit Function
      End If
      If eprob <= fprob Then
         pr = compbeta(eprob, n, i + 1#)
      Else
         pr = beta(fprob, i + 1#, n)
      End If
      If (pr > cprob) Then
         i = i + 1#
         If eprob <= fprob Then
            dfm = n - (n + i) * eprob
         Else
            dfm = (n + i) * fprob - i
         End If
         tpr = n / (n + i) * binomialTerm(i, n, fprob, eprob, dfm, 0#)
         If (pr < (1.00001) * tpr) Then
            While (tpr > cprob)
               i = i + 1#
               tpr = tpr * ((n + i - 1#) * fprob) / i
            Wend
         Else
            pr = pr - tpr
            If (pr <= cprob) Then
               critcompnegbinom = i
               Exit Function
            ElseIf (tpr < 0.000000000000001 * pr) Then
               If (tpr < cSmall) Then
                  critcompnegbinom = i
               Else
                  critcompnegbinom = i + Int((pr - cprob) / tpr)
               End If
               Exit Function
            End If
            Dim temp As Double, temp2 As Double
            temp = (pr - cprob) / tpr
            If (temp > 10#) Then
               temp = Int(temp + 0.5)
               i = i + temp
               If eprob <= fprob Then
                  dfm = n - (n + i) * eprob
               Else
                  dfm = (n + i) * fprob - i
               End If
               temp2 = n / (n + i) * binomialTerm(i, n, fprob, eprob, dfm, 0#)
               i = i + temp * (tpr - temp2) / (2# * temp2)
            Else
               i = i + 1#
               tpr = tpr * ((n + i - 1#) * fprob) / i
               pr = pr - tpr
               If (pr <= cprob) Then
                  critcompnegbinom = i
                  Exit Function
               End If
               temp2 = (pr - cprob) / tpr
               If (temp2 < temp - 0.9) Then
                  While (pr > cprob)
                     i = i + 1#
                     tpr = tpr * ((n + i - 1#) * fprob) / i
                     pr = pr - tpr
                  Wend
                  critcompnegbinom = i
                  Exit Function
               Else
                  temp = Int(Log(cprob / pr) / Log(((n + i - 1#) * fprob) / i) + 0.5)
                  i = i + temp
                  If eprob <= fprob Then
                     dfm = n - (n + i) * eprob
                  Else
                     dfm = (n + i) * fprob - i
                  End If
                  temp2 = n / (n + i) * binomialTerm(i, n, fprob, eprob, dfm, 0#)
                  If (temp2 > nearly_zero) Then
                     temp = Log((cprob / pr) * (tpr / temp2)) / Log(((n + i - 1#) * fprob) / i)
                     i = i + temp
                  End If
               End If
            End If
         End If
      Else
         If eprob <= fprob Then
            dfm = n - (n + i) * eprob
         Else
            dfm = (n + i) * fprob - i
         End If
         tpr = n / (n + i) * binomialTerm(i, n, fprob, eprob, dfm, 0#)
         If (tpr < 0.000000000000001 * pr) Then
            If (tpr < cSmall) Then
               critcompnegbinom = i
            Else
               critcompnegbinom = i - Int((cprob - pr) / tpr)
            End If
            Exit Function
         End If
         While ((tpr < cSmall) And (pr <= cprob))
            pr = pr + tpr
            i = i - 1#
            If eprob <= fprob Then
               dfm = n - (n + i) * eprob
            Else
               dfm = (n + i) * fprob - i
            End If
            tpr = n / (n + i) * binomialTerm(i, n, fprob, eprob, dfm, 0#)
         Wend
         While (pr <= cprob)
            pr = pr + tpr
            i = i - 1#
            If i < 0# Then
               critcompnegbinom = 0#
               Exit Function
            End If
            tpr = tpr * (i + 1#) / ((n + i) * fprob)
         Wend
         critcompnegbinom = i + 1#
         Exit Function
      End If
   Wend
End Function

Private Function critneghyperg(ByVal j As Double, ByVal k As Double, ByVal m As Double, ByVal cprob As Double) As Double
'//i such that Pr(neghypergeometric(i,j,k,m)) >= cprob and  Pr(neghypergeometric(i-1,j,k,m)) < cprob
   Dim ha1 As Double, hprob As Double, hswap As Boolean
   If (cprob > 0.5) Then
      critneghyperg = critcompneghyperg(j, k, m, 1# - cprob)
      Exit Function
   End If
   Dim pr As Double, tpr As Double
   Dim i As Double, temp As Double, temp2 As Double, oneMinusP As Double
   pr = (m - k) / m
   i = invbeta(j * pr, pr * (k - j + 1#), cprob, oneMinusP) * (m - k)
   While (True)
      If (i < 0#) Then
         i = 0#
      ElseIf (i > m - k) Then
         i = m - k
      End If
      i = Int(i + 0.5)
      If (i >= max_crit) Then
         critneghyperg = i
         Exit Function
      End If
      pr = hypergeometric(i, j, m - k - i, k - j, False, ha1, hprob, hswap)
      tpr = 0#
      If (pr >= cprob) Then
         If (i = 0#) Then
            critneghyperg = 0#
            Exit Function
         End If
         tpr = hypergeometricTerm(j - 1#, i, k - j + 1#, m - k - i) * (k - j + 1#) / (m - j - i + 1#)
         If (pr < (1# + 0.00001) * tpr) Then
            tpr = tpr * ((i + 1#) * (m - j - i)) / ((j + i) * (m - i - k))
            i = i - 1#
            While (tpr > cprob)
               tpr = tpr * ((i + 1#) * (m - j - i)) / ((j + i) * (m - i - k))
               i = i - 1#
            Wend
         Else
            pr = pr - tpr
            If (pr < cprob) Then
               critneghyperg = i
               Exit Function
            End If
            i = i - 1#

            If (i = 0#) Then
               critneghyperg = 0#
               Exit Function
            End If
            temp = (pr - cprob) / tpr
            If (temp > 10) Then
               temp = Int(temp + 0.5)
               i = i - temp
               temp2 = hypergeometricTerm(j - 1#, i, k - j + 1#, m - k - i) * (k - j + 1#) / (m - j - i + 1#)
               i = i - temp * (tpr - temp2) / (2 * temp2)
            Else
               tpr = tpr * ((i + 1#) * (m - j - i)) / ((j + i) * (m - i - k))
               pr = pr - tpr
               If (pr < cprob) Then
                  critneghyperg = i
                  Exit Function
               End If
               i = i - 1#
               temp2 = (pr - cprob) / tpr
               If (temp2 < temp - 0.9) Then
                  While (pr >= cprob)
                     tpr = tpr * ((i + 1#) * (m - j - i)) / ((j + i) * (m - i - k))
                     pr = pr - tpr
                     i = i - 1#
                  Wend
                  critneghyperg = i + 1#
                  Exit Function
               Else
                  temp = Int(Log(cprob / pr) / Log(((i + 1#) * (m - j - i)) / ((j + i) * (m - i - k))) + 0.5)
                  i = i - temp
                  temp2 = hypergeometricTerm(j - 1#, i, k - j + 1#, m - k - i) * (k - j + 1#) / (m - j - i + 1#)
                  If (temp2 > nearly_zero) Then
                     temp = Log((cprob / pr) * (tpr / temp2)) / Log(((i + 1#) * (m - j - i)) / ((j + i) * (m - i - k)))
                     i = i - temp
                  End If
               End If
            End If
         End If
      Else
         While ((tpr < cSmall) And (pr < cprob))
            i = i + 1#
            tpr = hypergeometricTerm(j - 1#, i, k - j + 1#, m - k - i) * (k - j + 1#) / (m - j - i + 1#)
            pr = pr + tpr
         Wend
         While (pr < cprob)
            i = i + 1#
            tpr = tpr * ((j + i - 1#) * (m - i - k + 1#)) / (i * (m - j - i + 1#))
            pr = pr + tpr
         Wend
         critneghyperg = i
         Exit Function
      End If
   Wend
End Function

Private Function critcompneghyperg(ByVal j As Double, ByVal k As Double, ByVal m As Double, ByVal cprob As Double) As Double
'//i such that 1-Pr(neghypergeometric(i,j,k,m)) > cprob and  1-Pr(neghypergeometric(i-1,j,k,m)) <= cprob
   Dim ha1 As Double, hprob As Double, hswap As Boolean
   If (cprob > 0.5) Then
      critcompneghyperg = critneghyperg(j, k, m, 1# - cprob)
      Exit Function
   End If
   Dim pr As Double, tpr As Double
   Dim i As Double, temp As Double, temp2 As Double, oneMinusP As Double
   pr = (m - k) / m
   i = invcompbeta(j * pr, pr * (k - j + 1#), cprob, oneMinusP) * (m - k)
   While (True)
      If (i < 0#) Then
         i = 0#
      ElseIf (i > m - k) Then
         i = m - k
      End If
      i = Int(i + 0.5)
      If (i >= max_crit) Then
         critcompneghyperg = i
         Exit Function
      End If
      pr = hypergeometric(i, j, m - k - i, k - j, True, ha1, hprob, hswap)
      tpr = 0#
      If (pr > cprob) Then
         i = i + 1#
         tpr = hypergeometricTerm(j - 1#, i, k - j + 1#, m - k - i) * (k - j + 1#) / (m - j - i + 1#)
         If (pr < (1.00001) * tpr) Then
            Do While (tpr > cprob)
               i = i + 1#
               temp = m - j - i + 1#
               If temp = 0# Then Exit Do
               tpr = tpr * ((j + i - 1#) * (m - i - k + 1#)) / (i * temp)
            Loop
         Else
            pr = pr - tpr
            If (pr <= cprob) Then
               critcompneghyperg = i
               Exit Function
            End If
            temp = (pr - cprob) / tpr
            If (temp > 10#) Then
               temp = Int(temp + 0.5)
               i = i + temp
               temp2 = hypergeometricTerm(j - 1#, i, k - j + 1#, m - k - i) * (k - j + 1#) / (m - j - i + 1#)
               i = i + temp * (tpr - temp2) / (2 * temp2)
            Else
               i = i + 1#
               tpr = tpr * ((j + i - 1#) * (m - i - k + 1#)) / (i * (m - j - i + 1#))
               pr = pr - tpr
               If (pr <= cprob) Then
                  critcompneghyperg = i
                  Exit Function
               End If
               temp2 = (pr - cprob) / tpr
               If (temp2 < temp - 0.9) Then
                  While (pr > cprob)
                     i = i + 1#
                     tpr = tpr * ((j + i - 1#) * (m - i - k + 1#)) / (i * (m - j - i + 1#))
                     pr = pr - tpr
                  Wend
                  critcompneghyperg = i
                  Exit Function
               Else
                  temp = Int(Log(cprob / pr) / Log(((j + i - 1#) * (m - i - k + 1#)) / (i * (m - j - i + 1#))) + 0.5)
                  i = i + temp
                  temp2 = hypergeometricTerm(j - 1#, i, k - j + 1#, m - k - i) * (k - j + 1#) / (m - j - i + 1#)
                  If (temp2 > nearly_zero) Then
                     temp = Log((cprob / pr) * (tpr / temp2)) / Log(((j + i - 1#) * (m - i - k + 1#)) / (i * (m - j - i + 1#)))
                     i = i + temp
                  End If
               End If
            End If
         End If
      Else
         While ((tpr < cSmall) And (pr <= cprob))
            tpr = hypergeometricTerm(j - 1#, i, k - j + 1#, m - k - i) * (k - j + 1#) / (m - j - i + 1#)
            pr = pr + tpr
            i = i - 1#
         Wend
         While (pr <= cprob)
            tpr = tpr * ((i + 1#) * (m - j - i)) / ((j + i) * (m - i - k))
            pr = pr + tpr
            i = i - 1#
         Wend
         critcompneghyperg = i + 1#
         Exit Function
      End If
   Wend
End Function

Private Function AlterForIntegralChecks_Others(ByVal value As Double) As Double
   If NonIntegralValuesAllowed_Others Then
      AlterForIntegralChecks_Others = Int(value)
   ElseIf value <> Int(value) Then
      AlterForIntegralChecks_Others = [#VALUE!]
   Else
      AlterForIntegralChecks_Others = value
   End If
End Function

Private Function AlterForIntegralChecks_df(ByVal value As Double) As Double
   If NonIntegralValuesAllowed_df Then
      AlterForIntegralChecks_df = value
   Else
      AlterForIntegralChecks_df = AlterForIntegralChecks_Others(value)
   End If
End Function

Private Function AlterForIntegralChecks_NB(ByVal value As Double) As Double
   If NonIntegralValuesAllowed_NB Then
      AlterForIntegralChecks_NB = value
   Else
      AlterForIntegralChecks_NB = AlterForIntegralChecks_Others(value)
   End If
End Function

Private Function GetRidOfMinusZeroes(ByVal x As Double) As Double
   If x = 0# Then
      GetRidOfMinusZeroes = 0#
   Else
      GetRidOfMinusZeroes = x
   End If
End Function

Public Function pmf_geometric(ByVal failures As Double, ByVal success_prob As Double) As Double
   failures = AlterForIntegralChecks_Others(failures)
   If (success_prob < 0# Or success_prob > 1#) Then
      pmf_geometric = [#VALUE!]
   ElseIf failures < 0# Then
      pmf_geometric = 0#
   ElseIf success_prob = 1# Then
      If failures = 0# Then
         pmf_geometric = 1#
      Else
         pmf_geometric = 0#
      End If
   Else
      pmf_geometric = success_prob * Exp(failures * log0(-success_prob))
   End If
   pmf_geometric = GetRidOfMinusZeroes(pmf_geometric)
End Function

Public Function cdf_geometric(ByVal failures As Double, ByVal success_prob As Double) As Double
   failures = Int(failures)
   If (success_prob < 0# Or success_prob > 1#) Then
      cdf_geometric = [#VALUE!]
   ElseIf failures < 0# Then
      cdf_geometric = 0#
   ElseIf success_prob = 1# Then
      If failures >= 0# Then
         cdf_geometric = 1#
      Else
         cdf_geometric = 0#
      End If
   Else
      cdf_geometric = -expm1((failures + 1#) * log0(-success_prob))
   End If
   cdf_geometric = GetRidOfMinusZeroes(cdf_geometric)
End Function

Public Function comp_cdf_geometric(ByVal failures As Double, ByVal success_prob As Double) As Double
   failures = Int(failures)
   If (success_prob < 0# Or success_prob > 1#) Then
      comp_cdf_geometric = [#VALUE!]
   ElseIf failures < 0# Then
      comp_cdf_geometric = 1#
   ElseIf success_prob = 1# Then
      If failures >= 0# Then
         comp_cdf_geometric = 0#
      Else
         comp_cdf_geometric = 1#
      End If
   Else
      comp_cdf_geometric = Exp((failures + 1#) * log0(-success_prob))
   End If
   comp_cdf_geometric = GetRidOfMinusZeroes(comp_cdf_geometric)
End Function

Public Function crit_geometric(ByVal success_prob As Double, ByVal crit_prob As Double) As Double
   If (success_prob <= 0# Or success_prob > 1# Or crit_prob < 0# Or crit_prob > 1#) Then
      crit_geometric = [#VALUE!]
   ElseIf (crit_prob = 0#) Then
      crit_geometric = [#VALUE!]
   ElseIf (success_prob = 1#) Then
      crit_geometric = 0#
   ElseIf (crit_prob = 1#) Then
      crit_geometric = [#VALUE!]
   Else
      crit_geometric = Int(log0(-crit_prob) / log0(-success_prob) - 1#)
      If -expm1((crit_geometric + 1#) * log0(-success_prob)) < crit_prob Then
         crit_geometric = crit_geometric + 1#
      End If
   End If
   crit_geometric = GetRidOfMinusZeroes(crit_geometric)
End Function

Public Function comp_crit_geometric(ByVal success_prob As Double, ByVal crit_prob As Double) As Double
   If (success_prob <= 0# Or success_prob > 1# Or crit_prob < 0# Or crit_prob > 1#) Then
      comp_crit_geometric = [#VALUE!]
   ElseIf (crit_prob = 1#) Then
      comp_crit_geometric = [#VALUE!]
   ElseIf (success_prob = 1#) Then
      comp_crit_geometric = 0#
   ElseIf (crit_prob = 0#) Then
      comp_crit_geometric = [#VALUE!]
   Else
      comp_crit_geometric = Int(Log(crit_prob) / log0(-success_prob) - 1#)
      If Exp((comp_crit_geometric + 1#) * log0(-success_prob)) > crit_prob Then
         comp_crit_geometric = comp_crit_geometric + 1#
      End If
   End If
   comp_crit_geometric = GetRidOfMinusZeroes(comp_crit_geometric)
End Function

Public Function lcb_geometric(ByVal failures As Double, ByVal prob As Double) As Double
   failures = AlterForIntegralChecks_Others(failures)
   If (prob < 0# Or prob > 1# Or failures < 0#) Then
      lcb_geometric = [#VALUE!]
   ElseIf (prob = 1#) Then
      lcb_geometric = 1#
   Else
      lcb_geometric = -expm1(log0(-prob) / (failures + 1#))
   End If
   lcb_geometric = GetRidOfMinusZeroes(lcb_geometric)
End Function

Public Function ucb_geometric(ByVal failures As Double, ByVal prob As Double) As Double
   failures = AlterForIntegralChecks_Others(failures)
   If (prob < 0# Or prob > 1# Or failures < 0#) Then
      ucb_geometric = [#VALUE!]
   ElseIf (prob = 0# Or failures = 0#) Then
      ucb_geometric = 1#
   ElseIf (prob = 1#) Then
      ucb_geometric = 0#
   Else
      ucb_geometric = -expm1(Log(prob) / failures)
   End If
   ucb_geometric = GetRidOfMinusZeroes(ucb_geometric)
End Function

Public Function pmf_negbinomial(ByVal failures As Double, ByVal success_prob As Double, ByVal successes_reqd As Double) As Double
   Dim q As Double, dfm As Double
   failures = AlterForIntegralChecks_Others(failures)
   successes_reqd = AlterForIntegralChecks_NB(successes_reqd)
   If (success_prob < 0# Or success_prob > 1# Or successes_reqd <= 0#) Then
      pmf_negbinomial = [#VALUE!]
   ElseIf (successes_reqd + failures > 0#) Then
      q = 1# - success_prob
      If success_prob <= q Then
         dfm = successes_reqd - (successes_reqd + failures) * success_prob
      Else
         dfm = (successes_reqd + failures) * q - failures
      End If
      pmf_negbinomial = successes_reqd / (successes_reqd + failures) * binomialTerm(failures, successes_reqd, q, success_prob, dfm, 0#)
   ElseIf (failures <> 0#) Then
      pmf_negbinomial = 0#
   Else
      pmf_negbinomial = 1#
   End If
   pmf_negbinomial = GetRidOfMinusZeroes(pmf_negbinomial)
End Function

Public Function cdf_negbinomial(ByVal failures As Double, ByVal success_prob As Double, ByVal successes_reqd As Double) As Double
   Dim q As Double, dfm As Double
   failures = Int(failures)
   successes_reqd = AlterForIntegralChecks_NB(successes_reqd)
   If (success_prob < 0# Or success_prob > 1# Or successes_reqd <= 0#) Then
      cdf_negbinomial = [#VALUE!]
   Else
      q = 1# - success_prob
      If q < success_prob Then
         cdf_negbinomial = compbeta(q, failures + 1, successes_reqd)
      Else
         cdf_negbinomial = beta(success_prob, successes_reqd, failures + 1)
      End If
   End If
   cdf_negbinomial = GetRidOfMinusZeroes(cdf_negbinomial)
End Function

Public Function comp_cdf_negbinomial(ByVal failures As Double, ByVal success_prob As Double, ByVal successes_reqd As Double) As Double
   Dim q As Double, dfm As Double
   failures = Int(failures)
   successes_reqd = AlterForIntegralChecks_NB(successes_reqd)
   If (success_prob < 0# Or success_prob > 1# Or successes_reqd <= 0#) Then
      comp_cdf_negbinomial = [#VALUE!]
   Else
      q = 1# - success_prob
      If q < success_prob Then
         comp_cdf_negbinomial = beta(q, failures + 1, successes_reqd)
      Else
         comp_cdf_negbinomial = compbeta(success_prob, successes_reqd, failures + 1)
      End If
   End If
   comp_cdf_negbinomial = GetRidOfMinusZeroes(comp_cdf_negbinomial)
End Function

Public Function crit_negbinomial(ByVal success_prob As Double, ByVal successes_reqd As Double, ByVal crit_prob As Double) As Double
   successes_reqd = AlterForIntegralChecks_NB(successes_reqd)
   If (success_prob <= 0# Or success_prob > 1# Or successes_reqd <= 0# Or crit_prob < 0# Or crit_prob > 1#) Then
      crit_negbinomial = [#VALUE!]
   ElseIf (crit_prob = 0#) Then
      crit_negbinomial = [#VALUE!]
   ElseIf (success_prob = 1#) Then
      crit_negbinomial = 0#
   ElseIf (crit_prob = 1#) Then
      crit_negbinomial = [#VALUE!]
   Else
      Dim i As Double, pr As Double
      crit_negbinomial = critnegbinom(successes_reqd, success_prob, 1# - success_prob, crit_prob)
      i = crit_negbinomial
      pr = cdf_negbinomial(i, success_prob, successes_reqd)
      If (pr = crit_prob) Then
      ElseIf (pr > crit_prob) Then
         i = i - 1#
         pr = cdf_negbinomial(i, success_prob, successes_reqd)
         If (pr >= crit_prob) Then
            crit_negbinomial = i
         End If
      Else
         crit_negbinomial = i + 1#
      End If
   End If
   crit_negbinomial = GetRidOfMinusZeroes(crit_negbinomial)
End Function

Public Function comp_crit_negbinomial(ByVal success_prob As Double, ByVal successes_reqd As Double, ByVal crit_prob As Double) As Double
   successes_reqd = AlterForIntegralChecks_NB(successes_reqd)
   If (success_prob <= 0# Or success_prob > 1# Or successes_reqd <= 0# Or crit_prob < 0# Or crit_prob > 1#) Then
      comp_crit_negbinomial = [#VALUE!]
   ElseIf (crit_prob = 1#) Then
      comp_crit_negbinomial = [#VALUE!]
   ElseIf (success_prob = 1#) Then
      comp_crit_negbinomial = 0#
   ElseIf (crit_prob = 0#) Then
      comp_crit_negbinomial = [#VALUE!]
   Else
      Dim i As Double, pr As Double
      comp_crit_negbinomial = critcompnegbinom(successes_reqd, success_prob, 1# - success_prob, crit_prob)
      i = comp_crit_negbinomial
      pr = comp_cdf_negbinomial(i, success_prob, successes_reqd)
      If (pr = crit_prob) Then
      ElseIf (pr < crit_prob) Then
         i = i - 1#
         pr = comp_cdf_negbinomial(i, success_prob, successes_reqd)
         If (pr <= crit_prob) Then
            comp_crit_negbinomial = i
         End If
      Else
         comp_crit_negbinomial = i + 1#
      End If
   End If
   comp_crit_negbinomial = GetRidOfMinusZeroes(comp_crit_negbinomial)
End Function

Public Function lcb_negbinomial(ByVal failures As Double, ByVal successes_reqd As Double, ByVal prob As Double) As Double
   failures = AlterForIntegralChecks_Others(failures)
   successes_reqd = AlterForIntegralChecks_NB(successes_reqd)
   If (prob < 0# Or prob > 1# Or failures < 0# Or successes_reqd <= 0#) Then
      lcb_negbinomial = [#VALUE!]
   ElseIf (prob = 0#) Then
      lcb_negbinomial = 0#
   ElseIf (prob = 1#) Then
      lcb_negbinomial = 1#
   Else
      Dim oneMinusP As Double
      lcb_negbinomial = invbeta(successes_reqd, failures + 1, prob, oneMinusP)
   End If
   lcb_negbinomial = GetRidOfMinusZeroes(lcb_negbinomial)
End Function

Public Function ucb_negbinomial(ByVal failures As Double, ByVal successes_reqd As Double, ByVal prob As Double) As Double
   failures = AlterForIntegralChecks_Others(failures)
   successes_reqd = AlterForIntegralChecks_NB(successes_reqd)
   If (prob < 0# Or prob > 1# Or failures < 0# Or successes_reqd <= 0#) Then
      ucb_negbinomial = [#VALUE!]
   ElseIf (prob = 0# Or failures = 0#) Then
      ucb_negbinomial = 1#
   ElseIf (prob = 1#) Then
      ucb_negbinomial = 0#
   Else
      Dim oneMinusP As Double
      ucb_negbinomial = invcompbeta(successes_reqd, failures, prob, oneMinusP)
   End If
   ucb_negbinomial = GetRidOfMinusZeroes(ucb_negbinomial)
End Function

Public Function pmf_binomial(ByVal sample_size As Double, ByVal successes As Double, ByVal success_prob As Double) As Double
   Dim q As Double, dfm As Double
   successes = AlterForIntegralChecks_Others(successes)
   sample_size = AlterForIntegralChecks_Others(sample_size)
   If (success_prob < 0# Or success_prob > 1# Or sample_size < 0#) Then
      pmf_binomial = [#VALUE!]
   Else
      q = 1# - success_prob
      If success_prob <= q Then
         dfm = sample_size * success_prob - successes
      Else
         dfm = (sample_size - successes) - sample_size * q
      End If
      pmf_binomial = binomialTerm(successes, sample_size - successes, success_prob, q, dfm, 0#)
   End If
   pmf_binomial = GetRidOfMinusZeroes(pmf_binomial)
End Function

Public Function cdf_binomial(ByVal sample_size As Double, ByVal successes As Double, ByVal success_prob As Double) As Double
   Dim q As Double, dfm As Double
   successes = Int(successes)
   sample_size = AlterForIntegralChecks_Others(sample_size)
   If (success_prob < 0# Or success_prob > 1# Or sample_size < 0#) Then
      cdf_binomial = [#VALUE!]
   Else
      q = 1# - success_prob
      If success_prob <= q Then
         dfm = sample_size * success_prob - successes
      Else
         dfm = (sample_size - successes) - sample_size * q
      End If
      cdf_binomial = binomial(successes, sample_size - successes, success_prob, q, dfm)
   End If
   cdf_binomial = GetRidOfMinusZeroes(cdf_binomial)
End Function

Public Function comp_cdf_binomial(ByVal sample_size As Double, ByVal successes As Double, ByVal success_prob As Double) As Double
   Dim q As Double, dfm As Double
   successes = Int(successes)
   sample_size = AlterForIntegralChecks_Others(sample_size)
   If (success_prob < 0# Or success_prob > 1# Or sample_size < 0#) Then
      comp_cdf_binomial = [#VALUE!]
   Else
      q = 1# - success_prob
      If success_prob <= q Then
         dfm = sample_size * success_prob - successes
      Else
         dfm = (sample_size - successes) - sample_size * q
      End If
      comp_cdf_binomial = compbinomial(successes, sample_size - successes, success_prob, q, dfm)
   End If
   comp_cdf_binomial = GetRidOfMinusZeroes(comp_cdf_binomial)
End Function

Public Function crit_binomial(ByVal sample_size As Double, ByVal success_prob As Double, ByVal crit_prob As Double) As Double
   sample_size = AlterForIntegralChecks_Others(sample_size)
   If (success_prob < 0# Or success_prob > 1# Or sample_size < 0# Or crit_prob < 0# Or crit_prob > 1#) Then
      crit_binomial = [#VALUE!]
   ElseIf (crit_prob = 0#) Then
      crit_binomial = [#VALUE!]
   ElseIf (success_prob = 0#) Then
      crit_binomial = 0#
   ElseIf (crit_prob = 1# Or success_prob = 1#) Then
      crit_binomial = sample_size
   Else
      Dim pr As Double, i As Double
      crit_binomial = critbinomial(sample_size, success_prob, crit_prob)
      i = crit_binomial
      pr = cdf_binomial(sample_size, i, success_prob)
      If (pr = crit_prob) Then
      ElseIf (pr > crit_prob) Then
         i = i - 1#
         pr = cdf_binomial(sample_size, i, success_prob)
         If (pr >= crit_prob) Then
            crit_binomial = i
         End If
      Else
         crit_binomial = i + 1#
      End If
   End If
   crit_binomial = GetRidOfMinusZeroes(crit_binomial)
End Function

Public Function comp_crit_binomial(ByVal sample_size As Double, ByVal success_prob As Double, ByVal crit_prob As Double) As Double
   sample_size = AlterForIntegralChecks_Others(sample_size)
   If (success_prob < 0# Or success_prob > 1# Or sample_size < 0# Or crit_prob < 0# Or crit_prob > 1#) Then
      comp_crit_binomial = [#VALUE!]
   ElseIf (crit_prob = 1#) Then
      comp_crit_binomial = [#VALUE!]
   ElseIf (crit_prob = 0# Or success_prob = 1#) Then
      comp_crit_binomial = sample_size
   ElseIf (success_prob = 0#) Then
      comp_crit_binomial = 0#
   Else
      Dim pr As Double, i As Double
      comp_crit_binomial = critcompbinomial(sample_size, success_prob, crit_prob)
      i = comp_crit_binomial
      pr = comp_cdf_binomial(sample_size, i, success_prob)
      If (pr = crit_prob) Then
      ElseIf (pr < crit_prob) Then
         i = i - 1#
         pr = comp_cdf_binomial(sample_size, i, success_prob)
         If (pr <= crit_prob) Then
            comp_crit_binomial = i
         End If
      Else
         comp_crit_binomial = i + 1#
      End If
   End If
   comp_crit_binomial = GetRidOfMinusZeroes(comp_crit_binomial)
End Function

Public Function lcb_binomial(ByVal sample_size As Double, ByVal successes As Double, ByVal prob As Double) As Double
   successes = AlterForIntegralChecks_Others(successes)
   sample_size = AlterForIntegralChecks_Others(sample_size)
   If (prob < 0# Or prob > 1#) Then
      lcb_binomial = [#VALUE!]
   ElseIf (sample_size < successes Or successes < 0#) Then
      lcb_binomial = [#VALUE!]
   ElseIf (prob = 0# Or successes = 0#) Then
      lcb_binomial = 0#
   ElseIf (prob = 1#) Then
      lcb_binomial = 1#
   Else
      Dim oneMinusP As Double
      lcb_binomial = invcompbinom(successes - 1#, sample_size - successes + 1#, prob, oneMinusP)
   End If
   lcb_binomial = GetRidOfMinusZeroes(lcb_binomial)
End Function

Public Function ucb_binomial(ByVal sample_size As Double, ByVal successes As Double, ByVal prob As Double) As Double
   successes = AlterForIntegralChecks_Others(successes)
   sample_size = AlterForIntegralChecks_Others(sample_size)
   If (prob < 0# Or prob > 1#) Then
      ucb_binomial = [#VALUE!]
   ElseIf (sample_size < successes Or successes < 0#) Then
      ucb_binomial = [#VALUE!]
   ElseIf (prob = 0# Or successes = sample_size#) Then
      ucb_binomial = 1#
   ElseIf (prob = 1#) Then
      ucb_binomial = 0#
   Else
      Dim oneMinusP As Double
      ucb_binomial = invbinom(successes, sample_size - successes, prob, oneMinusP)
   End If
   ucb_binomial = GetRidOfMinusZeroes(ucb_binomial)
End Function

Public Function pmf_poisson(ByVal mean As Double, ByVal i As Double) As Double
   i = AlterForIntegralChecks_Others(i)
   If (mean < 0#) Then
      pmf_poisson = [#VALUE!]
   ElseIf (i < 0#) Then
      pmf_poisson = 0#
   Else
      pmf_poisson = poissonTerm(i, mean, mean - i, 0#)
   End If
   pmf_poisson = GetRidOfMinusZeroes(pmf_poisson)
End Function

Public Function cdf_poisson(ByVal mean As Double, ByVal i As Double) As Double
   i = Int(i)
   If (mean < 0#) Then
      cdf_poisson = [#VALUE!]
   ElseIf (i < 0#) Then
      cdf_poisson = 0#
   Else
      cdf_poisson = cpoisson(i, mean, mean - i)
   End If
   cdf_poisson = GetRidOfMinusZeroes(cdf_poisson)
End Function

Public Function comp_cdf_poisson(ByVal mean As Double, ByVal i As Double) As Double
   i = Int(i)
   If (mean < 0#) Then
      comp_cdf_poisson = [#VALUE!]
   ElseIf (i < 0#) Then
      comp_cdf_poisson = 1#
   Else
      comp_cdf_poisson = comppoisson(i, mean, mean - i)
   End If
   comp_cdf_poisson = GetRidOfMinusZeroes(comp_cdf_poisson)
End Function

Public Function crit_poisson(ByVal mean As Double, ByVal crit_prob As Double) As Double
   If (crit_prob < 0# Or crit_prob > 1# Or mean < 0#) Then
      crit_poisson = [#VALUE!]
   ElseIf (crit_prob = 0#) Then
      crit_poisson = [#VALUE!]
   ElseIf (mean = 0#) Then
      crit_poisson = 0#
   ElseIf (crit_prob = 1#) Then
      crit_poisson = [#VALUE!]
   Else
      Dim pr As Double
      crit_poisson = critpoiss(mean, crit_prob)
      pr = cpoisson(crit_poisson, mean, mean - crit_poisson)
      If (pr = crit_prob) Then
      ElseIf (pr > crit_prob) Then
         crit_poisson = crit_poisson - 1#
         pr = cpoisson(crit_poisson, mean, mean - crit_poisson)
         If (pr < crit_prob) Then
            crit_poisson = crit_poisson + 1#
         End If
      Else
         crit_poisson = crit_poisson + 1#
      End If
   End If
   crit_poisson = GetRidOfMinusZeroes(crit_poisson)
End Function

Public Function comp_crit_poisson(ByVal mean As Double, ByVal crit_prob As Double) As Double
   If (crit_prob < 0# Or crit_prob > 1# Or mean < 0#) Then
      comp_crit_poisson = [#VALUE!]
   ElseIf (crit_prob = 1#) Then
      comp_crit_poisson = [#VALUE!]
   ElseIf (mean = 0#) Then
      comp_crit_poisson = 0#
   ElseIf (crit_prob = 0#) Then
      comp_crit_poisson = [#VALUE!]
   Else
      Dim pr As Double
      comp_crit_poisson = critcomppoiss(mean, crit_prob)
      pr = comppoisson(comp_crit_poisson, mean, mean - comp_crit_poisson)
      If (pr = crit_prob) Then
      ElseIf (pr < crit_prob) Then
         comp_crit_poisson = comp_crit_poisson - 1#
         pr = comppoisson(comp_crit_poisson, mean, mean - comp_crit_poisson)
         If (pr > crit_prob) Then
            comp_crit_poisson = comp_crit_poisson + 1#
         End If
      Else
         comp_crit_poisson = comp_crit_poisson + 1#
      End If
   End If
   comp_crit_poisson = GetRidOfMinusZeroes(comp_crit_poisson)
End Function

Public Function lcb_poisson(ByVal i As Double, ByVal prob As Double) As Double
   i = AlterForIntegralChecks_Others(i)
   If (prob < 0# Or prob > 1# Or i < 0#) Then
      lcb_poisson = [#VALUE!]
   ElseIf (prob = 0# Or i = 0#) Then
      lcb_poisson = 0#
   ElseIf (prob = 1#) Then
      lcb_poisson = [#VALUE!]
   Else
      lcb_poisson = invcomppoisson(i - 1#, prob)
   End If
   lcb_poisson = GetRidOfMinusZeroes(lcb_poisson)
End Function

Public Function ucb_poisson(ByVal i As Double, ByVal prob As Double) As Double
   i = AlterForIntegralChecks_Others(i)
   If (prob <= 0# Or prob > 1#) Then
      ucb_poisson = [#VALUE!]
   ElseIf (i < 0#) Then
      ucb_poisson = [#VALUE!]
   ElseIf (prob = 1#) Then
      ucb_poisson = 0#
   Else
      ucb_poisson = invpoisson(i, prob)
   End If
   ucb_poisson = GetRidOfMinusZeroes(ucb_poisson)
End Function

Public Function pmf_hypergeometric(ByVal type1s As Double, ByVal sample_size As Double, ByVal tot_type1 As Double, ByVal pop_size As Double) As Double
   type1s = AlterForIntegralChecks_Others(type1s)
   sample_size = AlterForIntegralChecks_Others(sample_size)
   tot_type1 = AlterForIntegralChecks_Others(tot_type1)
   pop_size = AlterForIntegralChecks_Others(pop_size)
   If (sample_size < 0# Or tot_type1 < 0# Or sample_size > pop_size Or tot_type1 > pop_size) Then
      pmf_hypergeometric = [#VALUE!]
   Else
      pmf_hypergeometric = hypergeometricTerm(type1s, sample_size - type1s, tot_type1 - type1s, pop_size - tot_type1 - sample_size + type1s)
   End If
   pmf_hypergeometric = GetRidOfMinusZeroes(pmf_hypergeometric)
End Function

Public Function cdf_hypergeometric(ByVal type1s As Double, ByVal sample_size As Double, ByVal tot_type1 As Double, ByVal pop_size As Double) As Double
   type1s = Int(type1s)
   sample_size = AlterForIntegralChecks_Others(sample_size)
   tot_type1 = AlterForIntegralChecks_Others(tot_type1)
   pop_size = AlterForIntegralChecks_Others(pop_size)
   If (sample_size < 0# Or tot_type1 < 0# Or sample_size > pop_size Or tot_type1 > pop_size) Then
      cdf_hypergeometric = [#VALUE!]
   Else
      Dim ha1 As Double, hprob As Double, hswap As Boolean
      cdf_hypergeometric = hypergeometric(type1s, sample_size - type1s, tot_type1 - type1s, pop_size - tot_type1 - sample_size + type1s, False, ha1, hprob, hswap)
   End If
   cdf_hypergeometric = GetRidOfMinusZeroes(cdf_hypergeometric)
End Function

Public Function comp_cdf_hypergeometric(ByVal type1s As Double, ByVal sample_size As Double, ByVal tot_type1 As Double, ByVal pop_size As Double) As Double
   type1s = Int(type1s)
   sample_size = AlterForIntegralChecks_Others(sample_size)
   tot_type1 = AlterForIntegralChecks_Others(tot_type1)
   pop_size = AlterForIntegralChecks_Others(pop_size)
   If (sample_size < 0# Or tot_type1 < 0# Or sample_size > pop_size Or tot_type1 > pop_size) Then
      comp_cdf_hypergeometric = [#VALUE!]
   Else
      Dim ha1 As Double, hprob As Double, hswap As Boolean
      comp_cdf_hypergeometric = hypergeometric(type1s, sample_size - type1s, tot_type1 - type1s, pop_size - tot_type1 - sample_size + type1s, True, ha1, hprob, hswap)
   End If
   comp_cdf_hypergeometric = GetRidOfMinusZeroes(comp_cdf_hypergeometric)
End Function

Public Function crit_hypergeometric(ByVal sample_size As Double, ByVal tot_type1 As Double, ByVal pop_size As Double, ByVal crit_prob As Double) As Double
   sample_size = AlterForIntegralChecks_Others(sample_size)
   tot_type1 = AlterForIntegralChecks_Others(tot_type1)
   pop_size = AlterForIntegralChecks_Others(pop_size)
   If (crit_prob < 0# Or crit_prob > 1#) Then
      crit_hypergeometric = [#VALUE!]
   ElseIf (sample_size < 0# Or tot_type1 < 0# Or sample_size > pop_size Or tot_type1 > pop_size) Then
      crit_hypergeometric = [#VALUE!]
   ElseIf (crit_prob = 0#) Then
      crit_hypergeometric = [#VALUE!]
   ElseIf (sample_size = 0# Or tot_type1 = 0#) Then
      crit_hypergeometric = 0#
   ElseIf (sample_size = pop_size Or tot_type1 = pop_size) Then
      crit_hypergeometric = Min(sample_size, tot_type1)
   ElseIf (crit_prob = 1#) Then
      crit_hypergeometric = Min(sample_size, tot_type1)
   Else
      Dim ha1 As Double, hprob As Double, hswap As Boolean
      Dim i As Double, pr As Double
      crit_hypergeometric = crithyperg(sample_size, tot_type1, pop_size, crit_prob)
      i = crit_hypergeometric
      pr = hypergeometric(i, sample_size - i, tot_type1 - i, pop_size - tot_type1 - sample_size + i, False, ha1, hprob, hswap)
      If (pr = crit_prob) Then
      ElseIf (pr > crit_prob) Then
         i = i - 1#
         pr = hypergeometric(i, sample_size - i, tot_type1 - i, pop_size - tot_type1 - sample_size + i, False, ha1, hprob, hswap)
         If (pr >= crit_prob) Then
            crit_hypergeometric = i
         End If
      Else
         crit_hypergeometric = i + 1#
      End If
   End If
   crit_hypergeometric = GetRidOfMinusZeroes(crit_hypergeometric)
End Function

Public Function comp_crit_hypergeometric(ByVal sample_size As Double, ByVal tot_type1 As Double, ByVal pop_size As Double, ByVal crit_prob As Double) As Double
   sample_size = AlterForIntegralChecks_Others(sample_size)
   tot_type1 = AlterForIntegralChecks_Others(tot_type1)
   pop_size = AlterForIntegralChecks_Others(pop_size)
   If (crit_prob < 0# Or crit_prob > 1#) Then
      comp_crit_hypergeometric = [#VALUE!]
   ElseIf (sample_size < 0# Or tot_type1 < 0# Or sample_size > pop_size Or tot_type1 > pop_size) Then
      comp_crit_hypergeometric = [#VALUE!]
   ElseIf (crit_prob = 1#) Then
      comp_crit_hypergeometric = [#VALUE!]
   ElseIf (sample_size = 0# Or tot_type1 = 0#) Then
      comp_crit_hypergeometric = 0#
   ElseIf (sample_size = pop_size Or tot_type1 = pop_size) Then
      comp_crit_hypergeometric = Min(sample_size, tot_type1)
   ElseIf (crit_prob = 0#) Then
      comp_crit_hypergeometric = Min(sample_size, tot_type1)
   Else
      Dim ha1 As Double, hprob As Double, hswap As Boolean
      Dim i As Double, pr As Double
      comp_crit_hypergeometric = critcomphyperg(sample_size, tot_type1, pop_size, crit_prob)
      i = comp_crit_hypergeometric
      pr = hypergeometric(i, sample_size - i, tot_type1 - i, pop_size - tot_type1 - sample_size + i, True, ha1, hprob, hswap)
      If (pr = crit_prob) Then
      ElseIf (pr < crit_prob) Then
         i = i - 1#
         pr = hypergeometric(i, sample_size - i, tot_type1 - i, pop_size - tot_type1 - sample_size + i, True, ha1, hprob, hswap)
         If (pr <= crit_prob) Then
            comp_crit_hypergeometric = i
         End If
      Else
         comp_crit_hypergeometric = i + 1#
      End If
   End If
   comp_crit_hypergeometric = GetRidOfMinusZeroes(comp_crit_hypergeometric)
End Function

Public Function lcb_hypergeometric(ByVal type1s As Double, ByVal sample_size As Double, ByVal pop_size As Double, ByVal prob As Double) As Double
   type1s = AlterForIntegralChecks_Others(type1s)
   sample_size = AlterForIntegralChecks_Others(sample_size)
   pop_size = AlterForIntegralChecks_Others(pop_size)
   If (prob < 0# Or prob > 1#) Then
      lcb_hypergeometric = [#VALUE!]
   ElseIf (type1s < 0# Or type1s > sample_size Or sample_size > pop_size) Then
      lcb_hypergeometric = [#VALUE!]
   ElseIf (prob = 0# Or type1s = 0# Or pop_size = sample_size) Then
      lcb_hypergeometric = type1s
   ElseIf (prob = 1#) Then
      lcb_hypergeometric = pop_size - (sample_size - type1s)
   ElseIf (prob < 0.5) Then
      lcb_hypergeometric = critneghyperg(type1s, sample_size, pop_size, prob * (1.000000000001)) + type1s
   Else
      lcb_hypergeometric = critcompneghyperg(type1s, sample_size, pop_size, (1# - prob) * (1# - 0.000000000001)) + type1s
   End If
   lcb_hypergeometric = GetRidOfMinusZeroes(lcb_hypergeometric)
End Function

Public Function ucb_hypergeometric(ByVal type1s As Double, ByVal sample_size As Double, ByVal pop_size As Double, ByVal prob As Double) As Double
   type1s = AlterForIntegralChecks_Others(type1s)
   sample_size = AlterForIntegralChecks_Others(sample_size)
   pop_size = AlterForIntegralChecks_Others(pop_size)
   If (prob < 0# Or prob > 1#) Then
      ucb_hypergeometric = [#VALUE!]
   ElseIf (type1s < 0# Or type1s > sample_size Or sample_size > pop_size) Then
      ucb_hypergeometric = [#VALUE!]
   ElseIf (prob = 0# Or type1s = sample_size Or pop_size = sample_size) Then
      ucb_hypergeometric = pop_size - (sample_size - type1s)
   ElseIf (prob = 1#) Then
      ucb_hypergeometric = type1s
   ElseIf (prob < 0.5) Then
      ucb_hypergeometric = critcompneghyperg(type1s + 1#, sample_size, pop_size, prob * (1# - 0.000000000001)) + type1s
   Else
      ucb_hypergeometric = critneghyperg(type1s + 1#, sample_size, pop_size, (1# - prob) * (1.000000000001)) + type1s
   End If
   ucb_hypergeometric = GetRidOfMinusZeroes(ucb_hypergeometric)
End Function

Public Function pmf_neghypergeometric(ByVal type2s As Double, ByVal type1s_reqd As Double, ByVal tot_type1 As Double, ByVal pop_size As Double) As Double
   type2s = AlterForIntegralChecks_Others(type2s)
   type1s_reqd = AlterForIntegralChecks_Others(type1s_reqd)
   tot_type1 = AlterForIntegralChecks_Others(tot_type1)
   pop_size = AlterForIntegralChecks_Others(pop_size)
   If (type1s_reqd <= 0# Or tot_type1 < type1s_reqd Or tot_type1 > pop_size) Then
      pmf_neghypergeometric = [#VALUE!]
   ElseIf (type2s < 0# Or tot_type1 + type2s > pop_size) Then
      If type2s = 0# Then
         pmf_neghypergeometric = 1#
      Else
         pmf_neghypergeometric = 0#
      End If
   Else
      pmf_neghypergeometric = hypergeometricTerm(type1s_reqd - 1#, type2s, tot_type1 - type1s_reqd + 1#, pop_size - tot_type1 - type2s) * (tot_type1 - type1s_reqd + 1#) / (pop_size - type1s_reqd - type2s + 1#)
   End If
   pmf_neghypergeometric = GetRidOfMinusZeroes(pmf_neghypergeometric)
End Function

Public Function cdf_neghypergeometric(ByVal type2s As Double, ByVal type1s_reqd As Double, ByVal tot_type1 As Double, ByVal pop_size As Double) As Double
   type2s = Int(type2s)
   type1s_reqd = AlterForIntegralChecks_Others(type1s_reqd)
   tot_type1 = AlterForIntegralChecks_Others(tot_type1)
   pop_size = AlterForIntegralChecks_Others(pop_size)
   If (type1s_reqd <= 0# Or tot_type1 < type1s_reqd Or tot_type1 > pop_size) Then
      cdf_neghypergeometric = [#VALUE!]
   ElseIf (tot_type1 + type2s > pop_size) Then
      cdf_neghypergeometric = 1#
   Else
      Dim ha1 As Double, hprob As Double, hswap As Boolean
      cdf_neghypergeometric = hypergeometric(type2s, type1s_reqd, pop_size - tot_type1 - type2s, tot_type1 - type1s_reqd, False, ha1, hprob, hswap)
   End If
   cdf_neghypergeometric = GetRidOfMinusZeroes(cdf_neghypergeometric)
End Function

Public Function comp_cdf_neghypergeometric(ByVal type2s As Double, ByVal type1s_reqd As Double, ByVal tot_type1 As Double, ByVal pop_size As Double) As Double
   type2s = Int(type2s)
   type1s_reqd = AlterForIntegralChecks_Others(type1s_reqd)
   tot_type1 = AlterForIntegralChecks_Others(tot_type1)
   pop_size = AlterForIntegralChecks_Others(pop_size)
   If (type1s_reqd <= 0# Or tot_type1 < type1s_reqd Or tot_type1 > pop_size) Then
      comp_cdf_neghypergeometric = [#VALUE!]
   ElseIf (tot_type1 + type2s > pop_size) Then
      comp_cdf_neghypergeometric = 0#
   Else
      Dim ha1 As Double, hprob As Double, hswap As Boolean
      comp_cdf_neghypergeometric = hypergeometric(type2s, type1s_reqd, pop_size - tot_type1 - type2s, tot_type1 - type1s_reqd, True, ha1, hprob, hswap)
   End If
   comp_cdf_neghypergeometric = GetRidOfMinusZeroes(comp_cdf_neghypergeometric)
End Function

Public Function crit_neghypergeometric(ByVal type1s_reqd As Double, ByVal tot_type1 As Double, ByVal pop_size As Double, ByVal crit_prob As Double) As Double
   type1s_reqd = AlterForIntegralChecks_Others(type1s_reqd)
   tot_type1 = AlterForIntegralChecks_Others(tot_type1)
   pop_size = AlterForIntegralChecks_Others(pop_size)
   If (crit_prob < 0# Or crit_prob > 1#) Then
      crit_neghypergeometric = [#VALUE!]
   ElseIf (type1s_reqd < 0# Or tot_type1 < type1s_reqd Or tot_type1 > pop_size) Then
      crit_neghypergeometric = [#VALUE!]
   ElseIf (crit_prob = 0#) Then
      crit_neghypergeometric = [#VALUE!]
   ElseIf (pop_size = tot_type1) Then
      crit_neghypergeometric = 0#
   ElseIf (crit_prob = 1#) Then
      crit_neghypergeometric = pop_size - tot_type1
   Else
      Dim ha1 As Double, hprob As Double, hswap As Boolean
      Dim i As Double, pr As Double
      crit_neghypergeometric = critneghyperg(type1s_reqd, tot_type1, pop_size, crit_prob)
      i = crit_neghypergeometric
      pr = hypergeometric(i, type1s_reqd, pop_size - tot_type1 - i, tot_type1 - type1s_reqd, False, ha1, hprob, hswap)
      If (pr = crit_prob) Then
      ElseIf (pr > crit_prob) Then
         i = i - 1#
         pr = hypergeometric(i, type1s_reqd, pop_size - tot_type1 - i, tot_type1 - type1s_reqd, False, ha1, hprob, hswap)
         If (pr >= crit_prob) Then
            crit_neghypergeometric = i
         End If
      Else
         crit_neghypergeometric = i + 1#
      End If
   End If
   crit_neghypergeometric = GetRidOfMinusZeroes(crit_neghypergeometric)
End Function

Public Function comp_crit_neghypergeometric(ByVal type1s_reqd As Double, ByVal tot_type1 As Double, ByVal pop_size As Double, ByVal crit_prob As Double) As Double
   type1s_reqd = AlterForIntegralChecks_Others(type1s_reqd)
   tot_type1 = AlterForIntegralChecks_Others(tot_type1)
   pop_size = AlterForIntegralChecks_Others(pop_size)
   If (crit_prob < 0# Or crit_prob > 1#) Then
      comp_crit_neghypergeometric = [#VALUE!]
   ElseIf (type1s_reqd <= 0# Or tot_type1 < type1s_reqd Or tot_type1 > pop_size) Then
      comp_crit_neghypergeometric = [#VALUE!]
   ElseIf (crit_prob = 1#) Then
      comp_crit_neghypergeometric = [#VALUE!]
   ElseIf (crit_prob = 0# Or pop_size = tot_type1) Then
      comp_crit_neghypergeometric = pop_size - tot_type1
   Else
      Dim ha1 As Double, hprob As Double, hswap As Boolean
      Dim i As Double, pr As Double
      comp_crit_neghypergeometric = critcompneghyperg(type1s_reqd, tot_type1, pop_size, crit_prob)
      i = comp_crit_neghypergeometric
      pr = hypergeometric(i, type1s_reqd, pop_size - tot_type1 - i, tot_type1 - type1s_reqd, True, ha1, hprob, hswap)
      If (pr = crit_prob) Then
      ElseIf (pr < crit_prob) Then
         i = i - 1#
         pr = hypergeometric(i, type1s_reqd, pop_size - tot_type1 - i, tot_type1 - type1s_reqd, True, ha1, hprob, hswap)
         If (pr <= crit_prob) Then
            comp_crit_neghypergeometric = i
         End If
      Else
         comp_crit_neghypergeometric = i + 1#
      End If
   End If
   comp_crit_neghypergeometric = GetRidOfMinusZeroes(comp_crit_neghypergeometric)
End Function

Public Function lcb_neghypergeometric(ByVal type2s As Double, ByVal type1s_reqd As Double, ByVal pop_size As Double, ByVal prob As Double) As Double
   type2s = AlterForIntegralChecks_Others(type2s)
   type1s_reqd = AlterForIntegralChecks_Others(type1s_reqd)
   pop_size = AlterForIntegralChecks_Others(pop_size)
   If (prob < 0# Or prob > 1#) Then
      lcb_neghypergeometric = [#VALUE!]
   ElseIf (type1s_reqd <= 0# Or type1s_reqd > pop_size Or type2s > pop_size - type1s_reqd) Then
      lcb_neghypergeometric = [#VALUE!]
   ElseIf (prob = 0# Or pop_size = type2s + type1s_reqd) Then
      lcb_neghypergeometric = type1s_reqd
   ElseIf (prob = 1#) Then
      lcb_neghypergeometric = pop_size - type2s
   ElseIf (prob < 0.5) Then
      lcb_neghypergeometric = critneghyperg(type1s_reqd, type2s + type1s_reqd, pop_size, prob * (1.000000000001)) + type1s_reqd
   Else
      lcb_neghypergeometric = critcompneghyperg(type1s_reqd, type2s + type1s_reqd, pop_size, (1# - prob) * (1# - 0.000000000001)) + type1s_reqd
   End If
   lcb_neghypergeometric = GetRidOfMinusZeroes(lcb_neghypergeometric)
End Function

Public Function ucb_neghypergeometric(ByVal type2s As Double, ByVal type1s_reqd As Double, ByVal pop_size As Double, ByVal prob As Double) As Double
   type2s = AlterForIntegralChecks_Others(type2s)
   type1s_reqd = AlterForIntegralChecks_Others(type1s_reqd)
   pop_size = AlterForIntegralChecks_Others(pop_size)
   If (prob < 0# Or prob > 1#) Then
      ucb_neghypergeometric = [#VALUE!]
   ElseIf (type1s_reqd <= 0# Or type1s_reqd > pop_size Or type2s > pop_size - type1s_reqd) Then
      ucb_neghypergeometric = [#VALUE!]
   ElseIf (prob = 0# Or type2s = 0# Or pop_size = type2s + type1s_reqd) Then
      ucb_neghypergeometric = pop_size - type2s
   ElseIf (prob = 1#) Then
      ucb_neghypergeometric = type1s_reqd
   ElseIf (prob < 0.5) Then
      ucb_neghypergeometric = critcompneghyperg(type1s_reqd, type2s + type1s_reqd - 1#, pop_size, prob * (1# - 0.000000000001)) + type1s_reqd - 1#
   Else
      ucb_neghypergeometric = critneghyperg(type1s_reqd, type2s + type1s_reqd - 1#, pop_size, (1# - prob) * (1.000000000001)) + type1s_reqd - 1#
   End If
   ucb_neghypergeometric = GetRidOfMinusZeroes(ucb_neghypergeometric)
End Function

Public Function pdf_triangular(ByVal x As Double, ByVal Min As Double, ByVal mode As Double, ByVal Max As Double) As Double
   If (Min > mode Or mode > Max) Then
      pdf_triangular = [#VALUE!]
   ElseIf (x <= Min Or x >= Max) Then
      pdf_triangular = 0#
   ElseIf (x <= mode) Then
      pdf_triangular = 2# * (x - Min) / (mode - Min) / (Max - Min)
   Else
      pdf_triangular = 2# * (Max - x) / (Max - mode) / (Max - Min)
   End If
End Function

Public Function cdf_triangular(ByVal x As Double, ByVal Min As Double, ByVal mode As Double, ByVal Max As Double) As Double
   If (Min > mode Or mode > Max) Then
      cdf_triangular = [#VALUE!]
   ElseIf (x <= Min) Then
      cdf_triangular = 0#
   ElseIf (x >= Max) Then
      cdf_triangular = 1#
   ElseIf (x <= mode) Then
      cdf_triangular = ((x - Min) / (mode - Min)) * ((x - Min) / (Max - Min))
   Else
      cdf_triangular = (mode - Min) / (Max - Min) + (1 + (Max - x) / (Max - mode)) * ((x - mode) / (Max - Min))
   End If
End Function

Public Function comp_cdf_triangular(ByVal x As Double, ByVal Min As Double, ByVal mode As Double, ByVal Max As Double) As Double
   If (Min > mode Or mode > Max) Then
      comp_cdf_triangular = [#VALUE!]
   ElseIf (x <= Min) Then
      comp_cdf_triangular = 1#
   ElseIf (x >= Max) Then
      comp_cdf_triangular = 0#
   ElseIf (x <= mode) Then
      comp_cdf_triangular = (Max - mode) / (Max - Min) + (1 + (x - Min) / (mode - Min)) * ((mode - x) / (Max - Min))
   Else
      comp_cdf_triangular = ((Max - x) / (Max - mode)) * ((Max - x) / (Max - Min))
   End If
End Function

Public Function inv_triangular(ByVal prob As Double, ByVal Min As Double, ByVal mode As Double, ByVal Max As Double) As Double
Dim temp As Double
   If (prob < 0# Or prob > 1# Or Min > mode Or mode > Max) Then
      inv_triangular = [#VALUE!]
   ElseIf (prob <= (mode - Min) / (Max - Min)) Then
      inv_triangular = Min + Sqr(prob) * Sqr(mode - Min) * Sqr(Max - Min)
   Else
      If prob > 0.5 Then
         inv_triangular = Max - Sqr(1# - prob) * Sqr(Max - Min) * Sqr(Max - mode)
      Else
         temp = (Max - mode) / (Max - Min)
         inv_triangular = mode + (Max - mode) * (prob - (mode - Min) / (Max - Min)) / (temp + Sqr(temp * (1# - prob)))
      End If
   End If
End Function

Public Function comp_inv_triangular(ByVal prob As Double, ByVal Min As Double, ByVal mode As Double, ByVal Max As Double) As Double
Dim temp As Double
   If (prob < 0# Or prob > 1# Or Min > mode Or mode > Max) Then
      comp_inv_triangular = [#VALUE!]
   ElseIf (prob <= (Max - mode) / (Max - Min)) Then
      comp_inv_triangular = Max - Sqr(prob) * Sqr(Max - mode) * Sqr(Max - Min)
   Else
      If prob > 0.5 Then
         comp_inv_triangular = Min + Sqr(1# - prob) * Sqr(mode - Min) * Sqr(Max - Min)
      Else
         temp = (mode - Min) / (Max - Min)
         comp_inv_triangular = mode - (mode - Min) * (prob - (Max - mode) / (Max - Min)) / (temp + Sqr(temp * (1# - prob)))
      End If
   End If
End Function

Public Function pdf_exponential(ByVal x As Double, ByVal lambda As Double) As Double
   If (lambda <= 0#) Then
      pdf_exponential = [#VALUE!]
   ElseIf (x < 0#) Then
      pdf_exponential = 0#
   Else
      pdf_exponential = Exp(-lambda * x + Log(lambda))
   End If
   pdf_exponential = GetRidOfMinusZeroes(pdf_exponential)
End Function

Public Function cdf_exponential(ByVal x As Double, ByVal lambda As Double) As Double
   If (lambda <= 0#) Then
      cdf_exponential = [#VALUE!]
   ElseIf (x < 0#) Then
      cdf_exponential = 0#
   Else
      cdf_exponential = -expm1(-lambda * x)
   End If
   cdf_exponential = GetRidOfMinusZeroes(cdf_exponential)
End Function

Public Function comp_cdf_exponential(ByVal x As Double, ByVal lambda As Double) As Double
   If (lambda <= 0#) Then
      comp_cdf_exponential = [#VALUE!]
   ElseIf (x < 0#) Then
      comp_cdf_exponential = 1#
   Else
      comp_cdf_exponential = Exp(-lambda * x)
   End If
   comp_cdf_exponential = GetRidOfMinusZeroes(comp_cdf_exponential)
End Function

Public Function inv_exponential(ByVal prob As Double, ByVal lambda As Double) As Double
   If (lambda <= 0# Or prob < 0# Or prob >= 1#) Then
      inv_exponential = [#VALUE!]
   Else
      inv_exponential = -log0(-prob) / lambda
   End If
   inv_exponential = GetRidOfMinusZeroes(inv_exponential)
End Function

Public Function comp_inv_exponential(ByVal prob As Double, ByVal lambda As Double) As Double
   If (lambda <= 0# Or prob <= 0# Or prob > 1#) Then
      comp_inv_exponential = [#VALUE!]
   Else
      comp_inv_exponential = -Log(prob) / lambda
   End If
   comp_inv_exponential = GetRidOfMinusZeroes(comp_inv_exponential)
End Function

Public Function pdf_normal(ByVal x As Double) As Double
   If (Abs(x) < 40#) Then
      pdf_normal = Exp(-x * x * 0.5) * OneOverSqrTwoPi
   Else
      pdf_normal = 0#
   End If
   pdf_normal = GetRidOfMinusZeroes(pdf_normal)
End Function

Public Function cdf_normal(ByVal x As Double) As Double
   cdf_normal = cnormal(x)
   cdf_normal = GetRidOfMinusZeroes(cdf_normal)
End Function

Public Function inv_normal(ByVal prob As Double) As Double
   If (prob <= 0# Or prob >= 1#) Then
      inv_normal = [#VALUE!]
   Else
      inv_normal = invcnormal(prob)
   End If
   inv_normal = GetRidOfMinusZeroes(inv_normal)
End Function

Public Function pdf_chi_sq(ByVal x As Double, ByVal df As Double) As Double
   df = AlterForIntegralChecks_df(df)
   pdf_chi_sq = pdf_gamma(x, df / 2#, 2#)
   pdf_chi_sq = GetRidOfMinusZeroes(pdf_chi_sq)
End Function

Public Function cdf_chi_sq(ByVal x As Double, ByVal df As Double) As Double
   df = AlterForIntegralChecks_df(df)
   If (df <= 0#) Then
      cdf_chi_sq = [#VALUE!]
   ElseIf (x <= 0#) Then
      cdf_chi_sq = 0#
   Else
      cdf_chi_sq = gamma(x / 2#, df / 2#)
   End If
   cdf_chi_sq = GetRidOfMinusZeroes(cdf_chi_sq)
End Function

Public Function comp_cdf_chi_sq(ByVal x As Double, ByVal df As Double) As Double
   df = AlterForIntegralChecks_df(df)
   If (df <= 0#) Then
      comp_cdf_chi_sq = [#VALUE!]
   ElseIf (x <= 0#) Then
      comp_cdf_chi_sq = 1#
   Else
      comp_cdf_chi_sq = compgamma(x / 2#, df / 2#)
   End If
   comp_cdf_chi_sq = GetRidOfMinusZeroes(comp_cdf_chi_sq)
End Function

Public Function inv_chi_sq(ByVal prob As Double, ByVal df As Double) As Double
   df = AlterForIntegralChecks_df(df)
   If (df <= 0# Or prob < 0# Or prob >= 1#) Then
      inv_chi_sq = [#VALUE!]
   ElseIf (prob = 0#) Then
      inv_chi_sq = 0#
   Else
      inv_chi_sq = 2# * invgamma(df / 2#, prob)
   End If
   inv_chi_sq = GetRidOfMinusZeroes(inv_chi_sq)
End Function

Public Function comp_inv_chi_sq(ByVal prob As Double, ByVal df As Double) As Double
   df = AlterForIntegralChecks_df(df)
   If (df <= 0# Or prob <= 0# Or prob > 1#) Then
      comp_inv_chi_sq = [#VALUE!]
   ElseIf (prob = 1#) Then
      comp_inv_chi_sq = 0#
   Else
      comp_inv_chi_sq = 2# * invcompgamma(df / 2#, prob)
   End If
   comp_inv_chi_sq = GetRidOfMinusZeroes(comp_inv_chi_sq)
End Function

Public Function pdf_gamma(ByVal x As Double, ByVal shape_param As Double, ByVal scale_param As Double) As Double
   Dim xs As Double
   If (shape_param <= 0# Or scale_param <= 0#) Then
      pdf_gamma = [#VALUE!]
   ElseIf (x < 0#) Then
      pdf_gamma = 0#
   ElseIf (x = 0#) Then
      If (shape_param < 1#) Then
         pdf_gamma = [#VALUE!]
      ElseIf (shape_param = 1#) Then
         pdf_gamma = 1# / scale_param
      Else
         pdf_gamma = 0#
      End If
   Else
      xs = x / scale_param
      pdf_gamma = poissonTerm(shape_param, xs, xs - shape_param, Log(shape_param) - Log(x))
   End If
   pdf_gamma = GetRidOfMinusZeroes(pdf_gamma)
End Function

Public Function cdf_gamma(ByVal x As Double, ByVal shape_param As Double, ByVal scale_param As Double) As Double
   If (shape_param <= 0# Or scale_param <= 0#) Then
      cdf_gamma = [#VALUE!]
   ElseIf (x <= 0#) Then
      cdf_gamma = 0#
   Else
      cdf_gamma = gamma(x / scale_param, shape_param)
   End If
   cdf_gamma = GetRidOfMinusZeroes(cdf_gamma)
End Function

Public Function comp_cdf_gamma(ByVal x As Double, ByVal shape_param As Double, ByVal scale_param As Double) As Double
   If (shape_param <= 0# Or scale_param <= 0#) Then
      comp_cdf_gamma = [#VALUE!]
   ElseIf (x <= 0#) Then
      comp_cdf_gamma = 1#
   Else
      comp_cdf_gamma = compgamma(x / scale_param, shape_param)
   End If
   comp_cdf_gamma = GetRidOfMinusZeroes(comp_cdf_gamma)
End Function

Public Function inv_gamma(ByVal prob As Double, ByVal shape_param As Double, ByVal scale_param As Double) As Double
   If (shape_param <= 0# Or scale_param <= 0# Or prob < 0# Or prob >= 1#) Then
      inv_gamma = [#VALUE!]
   ElseIf (prob = 0#) Then
      inv_gamma = 0#
   Else
      inv_gamma = scale_param * invgamma(shape_param, prob)
   End If
   inv_gamma = GetRidOfMinusZeroes(inv_gamma)
End Function

Public Function comp_inv_gamma(ByVal prob As Double, ByVal shape_param As Double, ByVal scale_param As Double) As Double
   If (shape_param <= 0# Or scale_param <= 0# Or prob <= 0# Or prob > 1#) Then
      comp_inv_gamma = [#VALUE!]
   ElseIf (prob = 1#) Then
      comp_inv_gamma = 0#
   Else
      comp_inv_gamma = scale_param * invcompgamma(shape_param, prob)
   End If
   comp_inv_gamma = GetRidOfMinusZeroes(comp_inv_gamma)
End Function

Private Function pdftdist(ByVal x As Double, ByVal k As Double) As Double
'//Probability density for a variate from t-distribution with k degress of freedom
   Dim a As Double, x2 As Double, k2 As Double, logterm As Double, c5 As Double
   If (k <= 0#) Then
      pdftdist = [#VALUE!]
   ElseIf (k > 1E+30) Then
      pdftdist = pdf_normal(x)
   Else
      If Abs(x) >= Min(1#, k) Then
         k2 = k / x
         x2 = x + k2
         k2 = k2 / x2
         x2 = x / x2
      Else
         x2 = x * x
         k2 = k + x2
         x2 = x2 / k2
         k2 = k / k2
      End If
      If (k2 < cSmall) Then
         logterm = Log(k) - 2# * Log(Abs(x))
      ElseIf (Abs(x2) < 0.5) Then
         logterm = log0(-x2)
      Else
         logterm = Log(k2)
      End If
      a = k * 0.5
      c5 = -1# / (k + 2#)
      pdftdist = Exp((a + 0.5) * logterm + a * log1(c5) - c5 + lfbaccdif1(0.5, a - 0.5)) * Sqr(a / ((1# + a))) * OneOverSqrTwoPi
   End If
End Function

Public Function pdf_tdist(ByVal x As Double, ByVal df As Double) As Double
   df = AlterForIntegralChecks_df(df)
   pdf_tdist = pdftdist(x, df)
   pdf_tdist = GetRidOfMinusZeroes(pdf_tdist)
End Function

Public Function cdf_tdist(ByVal x As Double, ByVal df As Double) As Double
   Dim tdistDensity As Double
   df = AlterForIntegralChecks_df(df)
   If (df <= 0#) Then
      cdf_tdist = [#VALUE!]
   Else
      cdf_tdist = tdist(x, df, tdistDensity)
   End If
   cdf_tdist = GetRidOfMinusZeroes(cdf_tdist)
End Function

Public Function inv_tdist(ByVal prob As Double, ByVal df As Double) As Double
   df = AlterForIntegralChecks_df(df)
   If (df <= 0#) Then
      inv_tdist = [#VALUE!]
   ElseIf (prob <= 0# Or prob >= 1#) Then
      inv_tdist = [#VALUE!]
   Else
      inv_tdist = invtdist(prob, df)
   End If
   inv_tdist = GetRidOfMinusZeroes(inv_tdist)
End Function

Public Function pdf_fdist(ByVal x As Double, ByVal df1 As Double, ByVal df2 As Double) As Double
   df1 = AlterForIntegralChecks_df(df1)
   df2 = AlterForIntegralChecks_df(df2)
   If (df1 <= 0# Or df2 <= 0#) Then
      pdf_fdist = [#VALUE!]
   ElseIf (x < 0#) Then
      pdf_fdist = 0#
   ElseIf (x = 0# And df1 > 2#) Then
      pdf_fdist = 0#
   ElseIf (x = 0# And df1 < 2#) Then
      pdf_fdist = [#VALUE!]
   ElseIf (x = 0#) Then
      pdf_fdist = 1#
   Else
      Dim p As Double, q As Double
      If x > 1# Then
         q = df2 / x
         p = q + df1
         q = q / p
         p = df1 / p
      Else
         p = df1 * x
         q = df2 + p
         p = p / q
         q = df2 / q
      End If
      'If p < cSmall And x <> 0# Or q < cSmall Then
      '   pdf_fdist = [#VALUE!]
      '   Exit Function
      'End If
      df2 = df2 / 2#
      df1 = df1 / 2#
      If (df1 >= 1#) Then
         df1 = df1 - 1#
         pdf_fdist = binomialTerm(df1, df2, p, q, df2 * p - df1 * q, Log((df1 + 1#) * q))
      Else
         pdf_fdist = df1 * df1 * q / (p * (df1 + df2)) * binomialTerm(df1, df2, p, q, df2 * p - df1 * q, 0#)
      End If
   End If
   pdf_fdist = GetRidOfMinusZeroes(pdf_fdist)
End Function

Public Function cdf_fdist(ByVal x As Double, ByVal df1 As Double, ByVal df2 As Double) As Double
   df1 = AlterForIntegralChecks_df(df1)
   df2 = AlterForIntegralChecks_df(df2)
   If (df1 <= 0# Or df2 <= 0#) Then
      cdf_fdist = [#VALUE!]
   ElseIf (x <= 0#) Then
      cdf_fdist = 0#
   Else
      Dim p As Double, q As Double
      If x > 1# Then
         q = df2 / x
         p = q + df1
         q = q / p
         p = df1 / p
      Else
         p = df1 * x
         q = df2 + p
         p = p / q
         q = df2 / q
      End If
      'If p < cSmall And x <> 0# Or q < cSmall Then
      '   cdf_fdist = [#VALUE!]
      '   Exit Function
      'End If
      df2 = df2 / 2#
      df1 = df1 / 2#
      If (p < 0.5) Then
          cdf_fdist = beta(p, df1, df2)
      Else
          cdf_fdist = compbeta(q, df2, df1)
      End If
   End If
   cdf_fdist = GetRidOfMinusZeroes(cdf_fdist)
End Function

Public Function comp_cdf_fdist(ByVal x As Double, ByVal df1 As Double, ByVal df2 As Double) As Double
   df1 = AlterForIntegralChecks_df(df1)
   df2 = AlterForIntegralChecks_df(df2)
   If (df1 <= 0# Or df2 <= 0#) Then
      comp_cdf_fdist = [#VALUE!]
   ElseIf (x <= 0#) Then
      comp_cdf_fdist = 1#
   Else
      Dim p As Double, q As Double
      If x > 1# Then
         q = df2 / x
         p = q + df1
         q = q / p
         p = df1 / p
      Else
         p = df1 * x
         q = df2 + p
         p = p / q
         q = df2 / q
      End If
      'If p < cSmall And x <> 0# Or q < cSmall Then
      '   comp_cdf_fdist = [#VALUE!]
      '   Exit Function
      'End If
      df2 = df2 / 2#
      df1 = df1 / 2#
      If (p < 0.5) Then
          comp_cdf_fdist = compbeta(p, df1, df2)
      Else
          comp_cdf_fdist = beta(q, df2, df1)
      End If
   End If
   comp_cdf_fdist = GetRidOfMinusZeroes(comp_cdf_fdist)
End Function

Public Function inv_fdist(ByVal prob As Double, ByVal df1 As Double, ByVal df2 As Double) As Double
   df1 = AlterForIntegralChecks_df(df1)
   df2 = AlterForIntegralChecks_df(df2)
   If (df1 <= 0# Or df2 <= 0# Or prob < 0# Or prob >= 1#) Then
      inv_fdist = [#VALUE!]
   ElseIf (prob = 0#) Then
      inv_fdist = 0#
   Else
      Dim temp As Double, oneMinusP As Double
      df1 = df1 / 2#
      df2 = df2 / 2#
      temp = invbeta(df1, df2, prob, oneMinusP)
      inv_fdist = df2 * temp / (df1 * oneMinusP)
      'If oneMinusP < cSmall Then inv_fdist = [#VALUE!]
   End If
   inv_fdist = GetRidOfMinusZeroes(inv_fdist)
End Function

Public Function comp_inv_fdist(ByVal prob As Double, ByVal df1 As Double, ByVal df2 As Double) As Double
   df1 = AlterForIntegralChecks_df(df1)
   df2 = AlterForIntegralChecks_df(df2)
   If (df1 <= 0# Or df2 <= 0# Or prob <= 0# Or prob > 1#) Then
      comp_inv_fdist = [#VALUE!]
   ElseIf (prob = 1#) Then
      comp_inv_fdist = 0#
   Else
      Dim temp As Double, oneMinusP As Double
      df1 = df1 / 2#
      df2 = df2 / 2#
      temp = invcompbeta(df1, df2, prob, oneMinusP)
      comp_inv_fdist = df2 * temp / (df1 * oneMinusP)
      'If oneMinusP < cSmall Then comp_inv_fdist = [#VALUE!]
   End If
   comp_inv_fdist = GetRidOfMinusZeroes(comp_inv_fdist)
End Function

Public Function pdf_beta(ByVal x As Double, ByVal shape_param1 As Double, ByVal shape_param2 As Double) As Double
   If (shape_param1 <= 0# Or shape_param2 <= 0#) Then
      pdf_beta = [#VALUE!]
   ElseIf (x < 0# Or x > 1#) Then
      pdf_beta = 0#
   ElseIf (x = 0# And shape_param1 < 1# Or x = 1# And shape_param2 < 1#) Then
      pdf_beta = [#VALUE!]
   ElseIf (x = 0# And shape_param1 = 1#) Then
      pdf_beta = shape_param2
   ElseIf (x = 1# And shape_param2 = 1#) Then
      pdf_beta = shape_param1
   ElseIf ((x = 0#) Or (x = 1#)) Then
      pdf_beta = 0#
   Else
      Dim mx As Double, mn As Double
      mx = Max(shape_param1, shape_param2)
      mn = Min(shape_param1, shape_param2)
      pdf_beta = (binomialTerm(shape_param1, shape_param2, x, 1# - x, (shape_param1 * (x - 1#) + x * shape_param2), 0#) * mx / (mn + mx)) * mn / (x * (1# - x))
   End If
   pdf_beta = GetRidOfMinusZeroes(pdf_beta)
End Function

Public Function cdf_beta(ByVal x As Double, ByVal shape_param1 As Double, ByVal shape_param2 As Double) As Double
   If (shape_param1 <= 0# Or shape_param2 <= 0#) Then
      cdf_beta = [#VALUE!]
   ElseIf (x <= 0#) Then
      cdf_beta = 0#
   ElseIf (x >= 1#) Then
      cdf_beta = 1#
   Else
      cdf_beta = beta(x, shape_param1, shape_param2)
   End If
   cdf_beta = GetRidOfMinusZeroes(cdf_beta)
End Function

Public Function comp_cdf_beta(ByVal x As Double, ByVal shape_param1 As Double, ByVal shape_param2 As Double) As Double
   If (shape_param1 <= 0# Or shape_param2 <= 0#) Then
      comp_cdf_beta = [#VALUE!]
   ElseIf (x <= 0#) Then
      comp_cdf_beta = 1#
   ElseIf (x >= 1#) Then
      comp_cdf_beta = 0#
   Else
      comp_cdf_beta = compbeta(x, shape_param1, shape_param2)
   End If
   comp_cdf_beta = GetRidOfMinusZeroes(comp_cdf_beta)
End Function

Public Function inv_beta(ByVal prob As Double, ByVal shape_param1 As Double, ByVal shape_param2 As Double) As Double
   If (shape_param1 <= 0# Or shape_param2 <= 0# Or prob < 0# Or prob > 1#) Then
      inv_beta = [#VALUE!]
   Else
      Dim oneMinusP As Double
      inv_beta = invbeta(shape_param1, shape_param2, prob, oneMinusP)
   End If
   inv_beta = GetRidOfMinusZeroes(inv_beta)
End Function

Public Function comp_inv_beta(ByVal prob As Double, ByVal shape_param1 As Double, ByVal shape_param2 As Double) As Double
   If (shape_param1 <= 0# Or shape_param2 <= 0# Or prob < 0# Or prob > 1#) Then
      comp_inv_beta = [#VALUE!]
   Else
      Dim oneMinusP As Double
      comp_inv_beta = invcompbeta(shape_param1, shape_param2, prob, oneMinusP)
   End If
   comp_inv_beta = GetRidOfMinusZeroes(comp_inv_beta)
End Function

Private Function gamma_nc1(ByVal x As Double, ByVal a As Double, ByVal nc As Double, ByRef nc_derivative As Double) As Double
   Dim aa As Double, bb As Double, nc_dtemp As Double
   Dim n As Double, p As Double, w As Double, s As Double, ps As Double
   Dim result As Double, term As Double, ptx As Double, ptnc As Double
   If a <= 1# And x <= 1# Then
      n = a + Sqr(a ^ 2 + 4# * nc * x)
      If n > 0# Then n = Int(2# * nc * x / n)
   ElseIf a > x Then
      n = x / a
      n = Int(2# * nc * n / (1# + Sqr(1# + 4# * n * (nc / a))))
   ElseIf x >= a Then
      n = a / x
      n = Int(2# * nc / (n + Sqr(n ^ 2 + 4# * (nc / x))))
   Else
      Debug.Print x, a, nc
   End If
   aa = n + a
   bb = n
   ptnc = poissonTerm(n, nc, nc - n, 0#)
   ptx = poissonTerm(aa, x, x - aa, 0#)
   aa = aa + 1#
   bb = bb + 1#
   p = nc / bb
   ps = p
   nc_derivative = ps
   s = x / aa
   w = p
   term = s * w
   result = term
   If ptx > 0# Then
     While (((term > 0.000000000000001 * result) And (p > 1E-16 * w)) Or (ps > 1E-16 * nc_derivative))
       aa = aa + 1#
       bb = bb + 1#
       p = nc / bb * p
       ps = p * s
       nc_derivative = nc_derivative + ps
       s = x / aa * s
       w = w + p
       term = s * w
       result = result + term
     Wend
     w = w * ptnc
   Else
     w = comppoisson(n, nc, nc - n)
   End If
   gamma_nc1 = result * ptx * ptnc + comppoisson(a + bb, x, (x - a) - bb) * w
   ps = 1#
   nc_dtemp = 0#
   aa = n + a
   bb = n
   p = 1#
   s = ptx
   w = gamma(x, aa)
   term = p * w
   result = term
   While bb > 0# And ((term > 0.000000000000001 * result) Or (ps > 1E-16 * nc_dtemp))
       s = aa / x * s
       ps = p * s
       nc_dtemp = nc_dtemp + ps
       p = bb / nc * p
       w = w + s
       term = p * w
       result = result + term
       aa = aa - 1#
       bb = bb - 1#
   Wend
   If bb = 0# Then aa = a
   If n > 0# Then
      nc_dtemp = nc_derivative * ptx + nc_dtemp + p * aa / x * s
   Else
      nc_dtemp = poissonTerm(aa, x, x - aa, Log(nc_derivative * x + aa) - Log(x))
   End If
   gamma_nc1 = gamma_nc1 + result * ptnc + cpoisson(bb - 1#, nc, nc - bb + 1#) * w
   If nc_dtemp = 0# Then
      nc_derivative = 0#
   Else
      nc_derivative = poissonTerm(n, nc, nc - n, Log(nc_dtemp))
   End If
End Function

Private Function comp_gamma_nc1(ByVal x As Double, ByVal a As Double, ByVal nc As Double, ByRef nc_derivative As Double) As Double
   Dim aa As Double, bb As Double, nc_dtemp As Double
   Dim n As Double, p As Double, w As Double, s As Double, ps As Double
   Dim result As Double, term As Double, ptx As Double, ptnc As Double
   If a <= 1# And x <= 1# Then
      n = a + Sqr(a ^ 2 + 4# * nc * x)
      If n > 0# Then n = Int(2# * nc * x / n)
   ElseIf a > x Then
      n = x / a
      n = Int(2# * nc * n / (1# + Sqr(1# + 4# * n * (nc / a))))
   ElseIf x >= a Then
      n = a / x
      n = Int(2# * nc / (n + Sqr(n ^ 2 + 4# * (nc / x))))
   Else
      Debug.Print x, a, nc
   End If
   aa = n + a
   bb = n
   ptnc = poissonTerm(n, nc, nc - n, 0#)
   ptx = poissonTerm(aa, x, x - aa, 0#)
   s = 1#
   ps = 1#
   nc_dtemp = 0#
   p = 1#
   w = p
   term = 1#
   result = 0#
   If ptx > 0# Then
     While bb > 0# And (((term > 0.000000000000001 * result) And (p > 1E-16 * w)) Or (ps > 1E-16 * nc_dtemp))
      s = aa / x * s
      ps = p * s
      nc_dtemp = nc_dtemp + ps
      p = bb / nc * p
      term = s * w
      result = result + term
      w = w + p
      aa = aa - 1#
      bb = bb - 1#
     Wend
     w = w * ptnc
   Else
     w = cpoisson(n, nc, nc - n)
   End If
   If bb = 0# Then aa = a
   If n > 0# Then
      nc_dtemp = (nc_dtemp + p * aa / x * s) * ptx
   ElseIf aa = 0 And x > 0 Then
      nc_dtemp = 0#
   Else
      nc_dtemp = poissonTerm(aa, x, x - aa, Log(aa) - Log(x))
   End If
   comp_gamma_nc1 = result * ptx * ptnc + compgamma(x, aa) * w
   aa = n + a
   bb = n
   ps = 1#
   nc_derivative = 0#
   p = 1#
   s = ptx
   w = compgamma(x, aa)
   term = 0#
   result = term
   Do
       w = w + s
       aa = aa + 1#
       bb = bb + 1#
       p = nc / bb * p
       ps = p * s
       nc_derivative = nc_derivative + ps
       s = x / aa * s
       term = p * w
       result = result + term
   Loop While (((term > 0.000000000000001 * result) And (s > 1E-16 * w)) Or (ps > 1E-16 * nc_derivative))
   comp_gamma_nc1 = comp_gamma_nc1 + result * ptnc + comppoisson(bb, nc, nc - bb) * w
   nc_dtemp = nc_derivative + nc_dtemp
   If nc_dtemp = 0# Then
      nc_derivative = 0#
   Else
      nc_derivative = poissonTerm(n, nc, nc - n, Log(nc_dtemp))
   End If
End Function

Private Function inv_gamma_nc1(ByVal prob As Double, ByVal a As Double, ByVal nc As Double) As Double
'Uses approx in A&S 26.4.27 for to get initial estimate the modified NR to improve it.
Dim x As Double, pr As Double, dif As Double
Dim hi As Double, lo As Double, nc_derivative As Double
   If (prob > 0.5) Then
      inv_gamma_nc1 = comp_inv_gamma_nc1(1# - prob, a, nc)
      Exit Function
   End If

   lo = 0#
   hi = 1E+308
   pr = Exp(-nc)
   If pr > prob Then
      If 2# * prob > pr Then
         x = comp_inv_gamma((pr - prob) / pr, a + cSmall, 1#)
      Else
         x = inv_gamma(prob / pr, a + cSmall, 1#)
      End If
      If x < cSmall Then
         x = cSmall
         pr = gamma_nc1(x, a, nc, nc_derivative)
         If pr > prob Then
            inv_gamma_nc1 = 0#
            Exit Function
         End If
      End If
   Else
      x = inv_gamma(prob, (a + nc) / (1# + nc / (a + nc)), 1#)
      x = x * (1# + nc / (a + nc))
   End If
   dif = x
   Do
      pr = gamma_nc1(x, a, nc, nc_derivative)
      If pr < 3E-308 And nc_derivative = 0# Then
         lo = x
         dif = dif / 2#
         x = x - dif
      ElseIf nc_derivative = 0# Then
         hi = x
         dif = dif / 2#
         x = x - dif
      Else
         If pr < prob Then
            lo = x
         Else
            hi = x
         End If
         dif = -(pr / nc_derivative) * logdif(pr, prob)
         If x + dif < lo Then
            dif = (lo - x) / 2#
         ElseIf x + dif > hi Then
            dif = (hi - x) / 2#
         End If
         x = x + dif
      End If
   Loop While ((Abs(pr - prob) > prob * 0.00000000000001) And (Abs(dif) > Abs(x) * 0.0000000001))
   inv_gamma_nc1 = x
End Function

Private Function comp_inv_gamma_nc1(ByVal prob As Double, ByVal a As Double, ByVal nc As Double) As Double
'Uses approx in A&S 26.4.27 for to get initial estimate the modified NR to improve it.
Dim x As Double, pr As Double, dif As Double
Dim hi As Double, lo As Double, nc_derivative As Double
   If (prob > 0.5) Then
      comp_inv_gamma_nc1 = inv_gamma_nc1(1# - prob, a, nc)
      Exit Function
   End If

   lo = 0#
   hi = 1E+308
   pr = Exp(-nc)
   If pr > prob Then
      x = comp_inv_gamma(prob / pr, a + cSmall, 1#) ' Is this as small as x could be?
   Else
      x = comp_inv_gamma(prob, (a + nc) / (1# + nc / (a + nc)), 1#)
      x = x * (1# + nc / (a + nc))
   End If
   If x < cSmall Then x = cSmall
   dif = x
   Do
      pr = comp_gamma_nc1(x, a, nc, nc_derivative)
      If pr < 3E-308 And nc_derivative = 0# Then
         hi = x
         dif = dif / 2#
         x = x - dif
      ElseIf nc_derivative = 0# Then
         lo = x
         dif = dif / 2#
         x = x - dif
      Else
         If pr < prob Then
            hi = x
         Else
            lo = x
         End If
         dif = (pr / nc_derivative) * logdif(pr, prob)
         If x + dif < lo Then
            dif = (lo - x) / 2#
         ElseIf x + dif > hi Then
            dif = (hi - x) / 2#
         End If
         x = x + dif
      End If
   Loop While ((Abs(pr - prob) > prob * 0.00000000000001) And (Abs(dif) > Abs(x) * 0.0000000001))
   comp_inv_gamma_nc1 = x
End Function

Private Function ncp_gamma_nc1(ByVal prob As Double, ByVal x As Double, ByVal a As Double) As Double
'Uses Normal approx for difference of 2 poisson distributed variables  to get initial estimate the modified NR to improve it.
Dim ncp As Double, pr As Double, dif As Double, temp As Double, deriv As Double, b As Double, sqarg As Double, checked_nc_limit As Boolean, checked_0_limit As Boolean
Dim hi As Double, lo As Double
   If (prob > 0.5) Then
      ncp_gamma_nc1 = comp_ncp_gamma_nc1(1# - prob, x, a)
      Exit Function
   End If

   lo = 0#
   hi = nc_limit
   checked_0_limit = False
   checked_nc_limit = False
   temp = inv_normal(prob) ^ 2
   b = 2# * (x - a) + temp
   sqarg = b ^ 2 - 4 * ((x - a) ^ 2 - temp * x)
   If sqarg < 0 Then
      ncp = b / 2
   Else
      ncp = (b + Sqr(sqarg)) / 2
   End If
   ncp = Max(0#, Min(ncp, nc_limit))
   If ncp = 0# Then
      pr = cdf_gamma_nc(x, a, 0#)
      If pr < prob Then
         If (inv_gamma(prob, a, 1) <= x) Then
            ncp_gamma_nc1 = 0#
         Else
            ncp_gamma_nc1 = [#VALUE!]
         End If
         Exit Function
      Else
         checked_0_limit = True
      End If
   ElseIf ncp = nc_limit Then
      pr = cdf_gamma_nc(x, a, ncp)
      If pr > prob Then
         ncp_gamma_nc1 = [#VALUE!]
         Exit Function
      Else
         checked_nc_limit = True
      End If
   End If
   dif = ncp
   Do
      pr = cdf_gamma_nc(x, a, ncp)
      'Debug.Print ncp, pr, prob
      deriv = pdf_gamma_nc(x, a + 1#, ncp)
      If pr < 3E-308 And deriv = 0# Then
         hi = ncp
         dif = dif / 2#
         ncp = ncp - dif
      ElseIf deriv = 0# Then
         lo = ncp
         dif = dif / 2#
         ncp = ncp - dif
      Else
         If pr < prob Then
            hi = ncp
         Else
            lo = ncp
         End If
         dif = (pr / deriv) * logdif(pr, prob)
         If ncp + dif < lo Then
            dif = (lo - ncp) / 2#
            If Not checked_0_limit And (lo = 0#) Then
               temp = cdf_gamma_nc(x, a, lo)
               If temp < prob Then
                  If (inv_gamma(prob, a, 1) <= x) Then
                     ncp_gamma_nc1 = 0#
                  Else
                     ncp_gamma_nc1 = [#VALUE!]
                  End If
                  Exit Function
               Else
                  checked_0_limit = True
               End If
            End If
         ElseIf ncp + dif > hi Then
            dif = (hi - ncp) / 2#
            If Not checked_nc_limit And (hi = nc_limit) Then
               pr = cdf_gamma_nc(x, a, hi)
               If pr > prob Then
                  ncp_gamma_nc1 = [#VALUE!]
                  Exit Function
               Else
                  ncp = hi
                  deriv = pdf_gamma_nc(x, a + 1#, ncp)
                  dif = (pr / deriv) * logdif(pr, prob)
                  If ncp + dif < lo Then
                     dif = (lo - ncp) / 2#
                  End If
                  checked_nc_limit = True
               End If
            End If
         End If
         ncp = ncp + dif
      End If
   Loop While ((Abs(pr - prob) > prob * 0.00000000000001) And (Abs(dif) > Abs(ncp) * 0.0000000001))
   ncp_gamma_nc1 = ncp
   'Debug.Print "ncp_gamma_nc1", ncp_gamma_nc1
End Function

Private Function comp_ncp_gamma_nc1(ByVal prob As Double, ByVal x As Double, ByVal a As Double) As Double
'Uses Normal approx for difference of 2 poisson distributed variables  to get initial estimate the modified NR to improve it.
Dim ncp As Double, pr As Double, dif As Double, temp As Double, deriv As Double, b As Double, sqarg As Double, checked_nc_limit As Boolean, checked_0_limit As Boolean
Dim hi As Double, lo As Double
   If (prob > 0.5) Then
      comp_ncp_gamma_nc1 = ncp_gamma_nc1(1# - prob, x, a)
      Exit Function
   End If

   lo = 0#
   hi = nc_limit
   checked_0_limit = False
   checked_nc_limit = False
   temp = inv_normal(prob) ^ 2
   b = 2# * (x - a) + temp
   sqarg = b ^ 2 - 4 * ((x - a) ^ 2 - temp * x)
   If sqarg < 0 Then
      ncp = b / 2
   Else
      ncp = (b - Sqr(sqarg)) / 2
   End If
   ncp = Max(0#, ncp)
   If ncp <= 1# Then
      pr = comp_cdf_gamma_nc(x, a, 0#)
      If pr > prob Then
         If (comp_inv_gamma(prob, a, 1) <= x) Then
            comp_ncp_gamma_nc1 = 0#
         Else
            comp_ncp_gamma_nc1 = [#VALUE!]
         End If
         Exit Function
      Else
         checked_0_limit = True
      End If
      deriv = pdf_gamma_nc(x, a + 1#, ncp)
      If deriv = 0# Then
         ncp = nc_limit
      ElseIf a < 1 Then
         ncp = (prob - pr) / deriv
         If ncp >= nc_limit Then
            ncp = -(pr / deriv) * logdif(pr, prob)
         End If
      Else
         ncp = -(pr / deriv) * logdif(pr, prob)
      End If
   End If
   ncp = Min(ncp, nc_limit)
   If ncp = nc_limit Then
      pr = comp_cdf_gamma_nc(x, a, ncp)
      If pr < prob Then
         comp_ncp_gamma_nc1 = [#VALUE!]
         Exit Function
      Else
         deriv = pdf_gamma_nc(x, a + 1#, ncp)
         dif = -(pr / deriv) * logdif(pr, prob)
         If ncp + dif < lo Then
            dif = (lo - ncp) / 2#
         End If
         checked_nc_limit = True
      End If
   End If
   dif = ncp
   Do
      pr = comp_cdf_gamma_nc(x, a, ncp)
      'Debug.Print ncp, pr, prob
      deriv = pdf_gamma_nc(x, a + 1#, ncp)
      If pr < 3E-308 And deriv = 0# Then
         lo = ncp
         dif = dif / 2#
         ncp = ncp + dif
      ElseIf deriv = 0# Then
         hi = ncp
         dif = dif / 2#
         ncp = ncp - dif
      Else
         If pr < prob Then
            lo = ncp
         Else
            hi = ncp
         End If
         dif = -(pr / deriv) * logdif(pr, prob)
         If ncp + dif < lo Then
            dif = (lo - ncp) / 2#
            If Not checked_0_limit And (lo = 0#) Then
               temp = comp_cdf_gamma_nc(x, a, lo)
               If temp > prob Then
                  If (comp_inv_gamma(prob, a, 1) <= x) Then
                     comp_ncp_gamma_nc1 = 0#
                  Else
                     comp_ncp_gamma_nc1 = [#VALUE!]
                  End If
                  Exit Function
               Else
                  checked_0_limit = True
               End If
            End If
         ElseIf ncp + dif > hi Then
            If Not checked_nc_limit And (hi = nc_limit) Then
               ncp = hi
               pr = comp_cdf_gamma_nc(x, a, ncp)
               If pr < prob Then
                  comp_ncp_gamma_nc1 = [#VALUE!]
                  Exit Function
               Else
                  deriv = pdf_gamma_nc(x, a + 1#, ncp)
                  dif = -(pr / deriv) * logdif(pr, prob)
                  If ncp + dif < lo Then
                     dif = (lo - ncp) / 2#
                  End If
                  checked_nc_limit = True
               End If
            Else
               dif = (hi - ncp) / 2#
            End If
         End If
         ncp = ncp + dif
      End If
   Loop While ((Abs(pr - prob) > prob * 0.00000000000001) And (Abs(dif) > Abs(ncp) * 0.0000000001))
   comp_ncp_gamma_nc1 = ncp
   'Debug.Print "comp_ncp_gamma_nc1", comp_ncp_gamma_nc1
End Function

Public Function pdf_gamma_nc(ByVal x As Double, ByVal shape_param As Double, ByVal nc_param As Double) As Double
'// Calculate pdf of noncentral gamma
  Dim nc_derivative As Double
  If (shape_param < 0#) Or (nc_param < 0#) Or (nc_param > nc_limit) Then
     pdf_gamma_nc = [#VALUE!]
  ElseIf (x < 0#) Then
     pdf_gamma_nc = 0#
  ElseIf (shape_param = 0# And nc_param = 0# And x > 0#) Then
     pdf_gamma_nc = 0#
  ElseIf (x = 0# Or nc_param = 0#) Then
     pdf_gamma_nc = Exp(-nc_param) * pdf_gamma(x, shape_param, 1#)
  ElseIf shape_param >= 1# Then
     If x >= nc_param Then
        If (x < 1# Or x <= shape_param + nc_param) Then
           pdf_gamma_nc = gamma_nc1(x, shape_param, nc_param, nc_derivative)
        Else
           pdf_gamma_nc = comp_gamma_nc1(x, shape_param, nc_param, nc_derivative)
        End If
        pdf_gamma_nc = nc_derivative
     Else
        If (nc_param < 1# Or nc_param <= shape_param + x) Then
           pdf_gamma_nc = gamma_nc1(nc_param, shape_param, x, nc_derivative)
        Else
           pdf_gamma_nc = comp_gamma_nc1(nc_param, shape_param, x, nc_derivative)
        End If
        If nc_derivative = 0# Then
           pdf_gamma_nc = 0#
        Else
           pdf_gamma_nc = Exp(Log(nc_derivative) + (shape_param - 1#) * (Log(x) - Log(nc_param)))
        End If
     End If
  Else
     If x < nc_param Then
        If (x < 1# Or x <= shape_param + nc_param) Then
           pdf_gamma_nc = gamma_nc1(x, shape_param, nc_param, nc_derivative)
        Else
           pdf_gamma_nc = comp_gamma_nc1(x, shape_param, nc_param, nc_derivative)
        End If
        pdf_gamma_nc = nc_derivative
     Else
        If (nc_param < 1# Or nc_param <= shape_param + x) Then
           pdf_gamma_nc = gamma_nc1(nc_param, shape_param, x, nc_derivative)
        Else
           pdf_gamma_nc = comp_gamma_nc1(nc_param, shape_param, x, nc_derivative)
        End If
        If nc_derivative = 0# Then
           pdf_gamma_nc = 0#
        Else
           pdf_gamma_nc = Exp(Log(nc_derivative) + (shape_param - 1#) * (Log(x) - Log(nc_param)))
        End If
     End If
  End If
  pdf_gamma_nc = GetRidOfMinusZeroes(pdf_gamma_nc)
End Function

Public Function cdf_gamma_nc(ByVal x As Double, ByVal shape_param As Double, ByVal nc_param As Double) As Double
'// Calculate cdf of noncentral gamma
  Dim nc_derivative As Double
  If (shape_param < 0#) Or (nc_param < 0#) Or (nc_param > nc_limit) Then
     cdf_gamma_nc = [#VALUE!]
  ElseIf (x < 0#) Then
     cdf_gamma_nc = 0#
  ElseIf (x = 0# And shape_param = 0#) Then
     cdf_gamma_nc = Exp(-nc_param)
  ElseIf (shape_param + nc_param = 0#) Then    ' limit as shape_param+nc_param->0 is degenerate point mass at zero
     cdf_gamma_nc = 1#                         ' if fix central gamma, then works for degenerate poisson
  ElseIf (x = 0#) Then
     cdf_gamma_nc = 0#
  ElseIf (nc_param = 0#) Then
     cdf_gamma_nc = gamma(x, shape_param)
  'ElseIf (shape_param = 0#) Then              ' extends Ruben (1974) and Cohen (1988) recurrence
  '   cdf_gamma_nc = ((x + shape_param + 2#) * gamma_nc1(x, shape_param + 2#, nc_param) + (nc_param - shape_param - 2#) * gamma_nc1(x, shape_param + 4#, nc_param) - nc_param * gamma_nc1(x, shape_param + 6#, nc_param)) / x
  ElseIf (x < 1# Or x <= shape_param + nc_param) Then
     cdf_gamma_nc = gamma_nc1(x, shape_param, nc_param, nc_derivative)
  Else
     cdf_gamma_nc = 1# - comp_gamma_nc1(x, shape_param, nc_param, nc_derivative)
  End If
  cdf_gamma_nc = GetRidOfMinusZeroes(cdf_gamma_nc)
End Function

Public Function comp_cdf_gamma_nc(ByVal x As Double, ByVal shape_param As Double, ByVal nc_param As Double) As Double
'// Calculate 1-cdf of noncentral gamma
  Dim nc_derivative As Double
  If (shape_param < 0#) Or (nc_param < 0#) Or (nc_param > nc_limit) Then
     comp_cdf_gamma_nc = [#VALUE!]
  ElseIf (x < 0#) Then
     comp_cdf_gamma_nc = 1#
  ElseIf (x = 0# And shape_param = 0#) Then
     comp_cdf_gamma_nc = -expm1(-nc_param)
  ElseIf (shape_param + nc_param = 0#) Then     ' limit as shape_param+nc_param->0 is degenerate point mass at zero
     comp_cdf_gamma_nc = 0#                     ' if fix central gamma, then works for degenerate poisson
  ElseIf (x = 0#) Then
     comp_cdf_gamma_nc = 1
  ElseIf (nc_param = 0#) Then
     comp_cdf_gamma_nc = compgamma(x, shape_param)
  'ElseIf (shape_param = 0#) Then              ' extends Ruben (1974) and Cohen (1988) recurrence
  '   comp_cdf_gamma_nc = ((x + shape_param + 2#) * comp_gamma_nc1(x, shape_param + 2#, nc_param) + (nc_param - shape_param - 2#) * comp_gamma_nc1(x, shape_param + 4#, nc_param) - nc_param * comp_gamma_nc1(x, shape_param + 6#, nc_param)) / x
  ElseIf (x < 1# Or x >= shape_param + nc_param) Then
     comp_cdf_gamma_nc = comp_gamma_nc1(x, shape_param, nc_param, nc_derivative)
  Else
     comp_cdf_gamma_nc = 1# - gamma_nc1(x, shape_param, nc_param, nc_derivative)
  End If
  comp_cdf_gamma_nc = GetRidOfMinusZeroes(comp_cdf_gamma_nc)
End Function

Public Function inv_gamma_nc(ByVal prob As Double, ByVal shape_param As Double, ByVal nc_param As Double) As Double
   If (shape_param < 0# Or nc_param < 0# Or nc_param > nc_limit Or prob < 0# Or prob >= 1#) Then
      inv_gamma_nc = [#VALUE!]
   ElseIf (prob = 0# Or shape_param = 0# And prob <= Exp(-nc_param)) Then
      inv_gamma_nc = 0#
   Else
      inv_gamma_nc = inv_gamma_nc1(prob, shape_param, nc_param)
   End If
   inv_gamma_nc = GetRidOfMinusZeroes(inv_gamma_nc)
End Function

Public Function comp_inv_gamma_nc(ByVal prob As Double, ByVal shape_param As Double, ByVal nc_param As Double) As Double
   If (shape_param < 0# Or nc_param < 0# Or nc_param > nc_limit Or prob <= 0# Or prob > 1#) Then
      comp_inv_gamma_nc = [#VALUE!]
   ElseIf (prob = 1# Or shape_param = 0# And prob >= -expm1(-nc_param)) Then
      comp_inv_gamma_nc = 0#
   Else
      comp_inv_gamma_nc = comp_inv_gamma_nc1(prob, shape_param, nc_param)
   End If
   comp_inv_gamma_nc = GetRidOfMinusZeroes(comp_inv_gamma_nc)
End Function

Public Function ncp_gamma_nc(ByVal prob As Double, ByVal x As Double, ByVal shape_param As Double) As Double
   If (shape_param < 0# Or x < 0# Or prob <= 0# Or prob > 1#) Then
      ncp_gamma_nc = [#VALUE!]
   ElseIf (x = 0# And shape_param = 0#) Then
      ncp_gamma_nc = -Log(prob)
   ElseIf (shape_param = 0# And prob = 1#) Then
      ncp_gamma_nc = 0#
   ElseIf (x = 0# Or prob = 1#) Then
      ncp_gamma_nc = [#VALUE!]
   Else
      ncp_gamma_nc = ncp_gamma_nc1(prob, x, shape_param)
   End If
   ncp_gamma_nc = GetRidOfMinusZeroes(ncp_gamma_nc)
End Function

Public Function comp_ncp_gamma_nc(ByVal prob As Double, ByVal x As Double, ByVal shape_param As Double) As Double
   If (shape_param < 0# Or x < 0# Or prob < 0# Or prob >= 1#) Then
      comp_ncp_gamma_nc = [#VALUE!]
   ElseIf (x = 0# And shape_param = 0#) Then
      comp_ncp_gamma_nc = -log0(-prob)
   ElseIf (shape_param = 0# And prob = 0#) Then
      comp_ncp_gamma_nc = 0#
   ElseIf (x = 0# Or prob = 0#) Then
      comp_ncp_gamma_nc = [#VALUE!]
   Else
      comp_ncp_gamma_nc = comp_ncp_gamma_nc1(prob, x, shape_param)
   End If
   comp_ncp_gamma_nc = GetRidOfMinusZeroes(comp_ncp_gamma_nc)
End Function

Public Function pdf_Chi2_nc(ByVal x As Double, ByVal df As Double, ByVal nc As Double) As Double
'// Calculate pdf of noncentral chi-square
  df = AlterForIntegralChecks_df(df)
  pdf_Chi2_nc = 0.5 * pdf_gamma_nc(x / 2#, df / 2#, nc / 2#)
  pdf_Chi2_nc = GetRidOfMinusZeroes(pdf_Chi2_nc)
End Function

Public Function cdf_Chi2_nc(ByVal x As Double, ByVal df As Double, ByVal nc As Double) As Double
'// Calculate cdf of noncentral chi-square
'//   parametrized per Johnson & Kotz, SAS, etc. so that cdf_Chi2_nc(x,df,nc) = cdf_gamma_nc(x/2,df/2,nc/2)
'//   If Xi ~ N(Di,1) independent, then sum(Xi,i=1..n) ~ Chi2_nc(n,nc) with nc=sum(Di,i=1..n)
'//   Note that Knusel, Graybill, etc. use a different noncentrality parameter lambda=nc/2
  df = AlterForIntegralChecks_df(df)
  cdf_Chi2_nc = cdf_gamma_nc(x / 2#, df / 2#, nc / 2#)
  cdf_Chi2_nc = GetRidOfMinusZeroes(cdf_Chi2_nc)
End Function

Public Function comp_cdf_Chi2_nc(ByVal x As Double, ByVal df As Double, ByVal nc As Double) As Double
'// Calculate 1-cdf of noncentral chi-square
'//   parametrized per Johnson & Kotz, SAS, etc. so that cdf_Chi2_nc(x,df,nc) = cdf_gamma_nc(x/2,df/2,nc/2)
'//   If Xi ~ N(Di,1) independent, then sum(Xi,i=1..n) ~ Chi2_nc(n,nc) with nc=sum(Di,i=1..n)
'//   Note that Knusel, Graybill, etc. use a different noncentrality parameter lambda=nc/2
  df = AlterForIntegralChecks_df(df)
  comp_cdf_Chi2_nc = comp_cdf_gamma_nc(x / 2#, df / 2#, nc / 2#)
  comp_cdf_Chi2_nc = GetRidOfMinusZeroes(comp_cdf_Chi2_nc)
End Function

Public Function inv_Chi2_nc(ByVal prob As Double, ByVal df As Double, ByVal nc As Double) As Double
   df = AlterForIntegralChecks_df(df)
   inv_Chi2_nc = 2# * inv_gamma_nc(prob, df / 2#, nc / 2#)
   inv_Chi2_nc = GetRidOfMinusZeroes(inv_Chi2_nc)
End Function

Public Function comp_inv_Chi2_nc(ByVal prob As Double, ByVal df As Double, ByVal nc As Double) As Double
   df = AlterForIntegralChecks_df(df)
   comp_inv_Chi2_nc = 2# * comp_inv_gamma_nc(prob, df / 2#, nc / 2#)
   comp_inv_Chi2_nc = GetRidOfMinusZeroes(comp_inv_Chi2_nc)
End Function

Public Function ncp_Chi2_nc(ByVal prob As Double, ByVal x As Double, ByVal df As Double) As Double
   df = AlterForIntegralChecks_df(df)
   ncp_Chi2_nc = 2# * ncp_gamma_nc(prob, x / 2#, df / 2#)
   ncp_Chi2_nc = GetRidOfMinusZeroes(ncp_Chi2_nc)
End Function

Public Function comp_ncp_Chi2_nc(ByVal prob As Double, ByVal x As Double, ByVal df As Double) As Double
   df = AlterForIntegralChecks_df(df)
   comp_ncp_Chi2_nc = 2# * comp_ncp_gamma_nc(prob, x / 2#, df / 2#)
   comp_ncp_Chi2_nc = GetRidOfMinusZeroes(comp_ncp_Chi2_nc)
End Function

Private Function beta_nc1(ByVal x As Double, ByVal y As Double, ByVal a As Double, ByVal b As Double, ByVal nc As Double, ByRef nc_derivative As Double) As Double
'y is 1-x but held accurately to avoid possible cancellation errors
   Dim aa As Double, bb As Double, nc_dtemp As Double
   Dim n As Double, p As Double, w As Double, s As Double, ps As Double
   Dim result As Double, term As Double, ptx As Double, ptnc As Double
   bb = (x * nc - 1#) - a
   If bb < -1E+150 Then
      n = a / bb
      aa = n - nc * x * (n + b / bb)
      n = bb * (1# + Sqr(1 - (4# * aa / bb)))
      n = Int(2# * aa * (bb / n))
   Else
      aa = a - nc * x * (a + b)
      If (bb < 0#) Then
         n = bb - Sqr(bb ^ 2 - 4# * aa)
         n = Int(2# * aa / n)
      Else
         n = Int((bb + Sqr(bb ^ 2 - 4# * aa)) / 2#)
      End If
   End If
   If n < 0# Then
      n = 0#
   End If
   aa = n + a
   bb = n
   ptnc = poissonTerm(n, nc, nc - n, 0#)
   ptx = b * binomialTerm(aa, b, x, y, b * x - aa * y, 0#)  '  (aa + b)*(I(x, aa, b) - I(x, aa + 1, b))
   aa = aa + 1#
   bb = bb + 1#
   p = nc / bb
   ps = p
   nc_derivative = ps
   s = x / aa  ' (I(x, aa, b) - I(x, aa + 1, b)) / ptx
   w = p
   term = s * w
   result = term
   If ptx > 0 Then
     While (((term > 0.000000000000001 * result) And (p > 1E-16 * w)) Or (ps > 1E-16 * nc_derivative))
       s = (aa + b) * s
       aa = aa + 1#
       bb = bb + 1#
       p = nc / bb * p
       ps = p * s
       nc_derivative = nc_derivative + ps
       s = x / aa * s ' (I(x, aa, b) - I(x, aa + 1, b)) / ptx
       w = w + p
       term = s * w
       result = result + term
     Wend
     w = w * ptnc
   Else
     w = comppoisson(n, nc, nc - n)
   End If
   If x > y Then
      s = compbeta(y, b, a + (bb + 1#))
   Else
      s = beta(x, a + (bb + 1#), b)
   End If
   beta_nc1 = result * ptx * ptnc + s * w
   ps = 1#
   nc_dtemp = 0#
   aa = n + a
   bb = n
   p = 1#
   s = ptx / (aa + b) ' I(x, aa, b) - I(x, aa + 1, b)
   If x > y Then
      w = compbeta(y, b, aa) ' I(x, aa, b)
   Else
      w = beta(x, aa, b) ' I(x, aa, b)
   End If
   term = p * w
   result = term
   While bb > 0# And (((term > 0.000000000000001 * result) And (s > 1E-16 * w)) Or (ps > 1E-16 * nc_dtemp))
       s = aa / x * s
       ps = p * s
       nc_dtemp = nc_dtemp + ps
       p = bb / nc * p
       aa = aa - 1#
       bb = bb - 1#
       If bb = 0# Then aa = a
       s = s / (aa + b) ' I(x, aa, b) - I(x, aa + 1, b)
       w = w + s ' I(x, aa, b)
       term = p * w
       result = result + term
   Wend
   If n > 0# Then
      nc_dtemp = nc_derivative * ptx + nc_dtemp + p * aa / x * s
   ElseIf b = 0# Then
      nc_dtemp = 0#
   Else
      nc_dtemp = binomialTerm(aa, b, x, y, b * x - aa * y, Log(b) + Log((nc_derivative + aa / (x * (aa + b)))))
   End If
   nc_dtemp = nc_dtemp / y
   beta_nc1 = beta_nc1 + result * ptnc + cpoisson(bb - 1#, nc, nc - bb + 1#) * w
   If nc_dtemp = 0# Then
      nc_derivative = 0#
   Else
      nc_derivative = poissonTerm(n, nc, nc - n, Log(nc_dtemp))
   End If
End Function

Private Function comp_beta_nc1(ByVal x As Double, ByVal y As Double, ByVal a As Double, ByVal b As Double, ByVal nc As Double, ByRef nc_derivative As Double) As Double
'y is 1-x but held accurately to avoid possible cancellation errors
   Dim aa As Double, bb As Double, nc_dtemp As Double
   Dim n As Double, p As Double, w As Double, s As Double, ps As Double
   Dim result As Double, term As Double, ptx As Double, ptnc As Double
   bb = (x * nc - 1#) - a
   If bb < -1E+150 Then
      n = a / bb
      aa = n - nc * x * (n + b / bb)
      n = bb * (1# + Sqr(1 - (4# * aa / bb)))
      n = Int(2# * aa * (bb / n))
   Else
      aa = a - nc * x * (a + b)
      If (bb < 0#) Then
         n = bb - Sqr(bb ^ 2 - 4# * aa)
         n = Int(2# * aa / n)
      Else
         n = Int((bb + Sqr(bb ^ 2 - 4# * aa)) / 2#)
      End If
   End If
   If n < 0# Then
      n = 0#
   End If
   aa = n + a
   bb = n
   ptnc = poissonTerm(n, nc, nc - n, 0#)
   ptx = b / (aa + b) * binomialTerm(aa, b, x, y, b * x - aa * y, 0#) '(1 - I(x, aa + 1, b)) - (1 - I(x, aa, b))
   ps = 1#
   nc_dtemp = 0#
   p = 1#
   s = 1#
   w = p
   term = 1#
   result = 0#
   If ptx > 0 Then
     While bb > 0# And (((term > 0.000000000000001 * result) And (p > 1E-16 * w)) Or (ps > 1E-16 * nc_dtemp))
       s = aa / x * s
       ps = p * s
       nc_dtemp = nc_dtemp + ps
       p = bb / nc * p
       aa = aa - 1#
       bb = bb - 1#
       If bb = 0# Then aa = a
       s = s / (aa + b) ' (1 - I(x, aa + 1, b)) - (1 - I(x, aa + 1, b))
       term = s * w
       result = result + term
       w = w + p
     Wend
     w = w * ptnc
   Else
     w = cpoisson(n, nc, nc - n)
   End If
   If n > 0# Then
      nc_dtemp = (nc_dtemp + p * aa / x * s) * ptx
   ElseIf a = 0# Or b = 0# Then
      nc_dtemp = 0#
   Else
      nc_dtemp = binomialTerm(aa, b, x, y, b * x - aa * y, Log(b) + Log(aa / (x * (aa + b))))
   End If
   If x > y Then
      s = beta(y, b, aa)
   Else
      s = compbeta(x, aa, b)
   End If
   comp_beta_nc1 = result * ptx * ptnc + s * w
   aa = n + a
   bb = n
   p = 1#
   nc_derivative = 0#
   s = ptx
   If x > y Then
      w = beta(y, b, aa) '  1 - I(x, aa, b)
   Else
      w = compbeta(x, aa, b) ' 1 - I(x, aa, b)
   End If
   term = 0#
   result = term
   Do
       w = w + s ' 1 - I(x, aa, b)
       s = (aa + b) * s
       aa = aa + 1#
       bb = bb + 1#
       p = nc / bb * p
       ps = p * s
       nc_derivative = nc_derivative + ps
       s = x / aa * s ' (1 - I(x, aa + 1, b)) - (1 - I(x, aa, b))
       term = p * w
       result = result + term
   Loop While (((term > 0.000000000000001 * result) And (s > 1E-16 * w)) Or (ps > 1E-16 * nc_derivative))
   nc_dtemp = (nc_derivative + nc_dtemp) / y
   comp_beta_nc1 = comp_beta_nc1 + result * ptnc + comppoisson(bb, nc, nc - bb) * w
   If nc_dtemp = 0# Then
      nc_derivative = 0#
   Else
      nc_derivative = poissonTerm(n, nc, nc - n, Log(nc_dtemp))
   End If
End Function

Private Function inv_beta_nc1(ByVal prob As Double, ByVal a As Double, ByVal b As Double, ByVal nc As Double, ByRef oneMinusP As Double) As Double
'Uses approx in A&S 26.6.26 for to get initial estimate the modified NR to improve it.
Dim x As Double, y As Double, pr As Double, dif As Double, temp As Double
Dim hip As Double, lop As Double
Dim hix As Double, lox As Double, nc_derivative As Double
   If (prob > 0.5) Then
      inv_beta_nc1 = comp_inv_beta_nc1(1# - prob, a, b, nc, oneMinusP)
      Exit Function
   End If

   lop = 0#
   hip = 1#
   lox = 0#
   hix = 1#
   pr = Exp(-nc)
   If pr > prob Then
      If 2# * prob > pr Then
         x = invcompbeta(a + cSmall, b, (pr - prob) / pr, oneMinusP)
      Else
         x = invbeta(a + cSmall, b, prob / pr, oneMinusP)
      End If
      If x = 0# Then
         inv_beta_nc1 = 0#
         Exit Function
      Else
         temp = oneMinusP
         y = invbeta(a + nc ^ 2 / (a + 2 * nc), b, prob, oneMinusP)
         oneMinusP = (a + nc) * oneMinusP / (a + nc * (1# + y))
         If temp > oneMinusP Then
            oneMinusP = temp
         Else
            x = (a + 2# * nc) * y / (a + nc * (1# + y))
         End If
      End If
   Else
      y = invbeta(a + nc ^ 2 / (a + 2 * nc), b, prob, oneMinusP)
      x = (a + 2# * nc) * y / (a + nc * (1# + y))
      oneMinusP = (a + nc) * oneMinusP / (a + nc * (1# + y))
      If oneMinusP < cSmall Then
         oneMinusP = cSmall
         pr = beta_nc1(x, oneMinusP, a, b, nc, nc_derivative)
         If pr < prob Then
            inv_beta_nc1 = 1#
            oneMinusP = 0#
            Exit Function
         End If
      End If
   End If
   Do
      pr = beta_nc1(x, oneMinusP, a, b, nc, nc_derivative)
      If pr < 3E-308 And nc_derivative = 0# Then
         hip = oneMinusP
         lox = x
         dif = dif / 2#
         x = x - dif
         oneMinusP = oneMinusP + dif
      ElseIf nc_derivative = 0# Then
         lop = oneMinusP
         hix = x
         dif = dif / 2#
         x = x - dif
         oneMinusP = oneMinusP + dif
      Else
         If pr < prob Then
            hip = oneMinusP
            lox = x
         Else
            lop = oneMinusP
            hix = x
         End If
         dif = -(pr / nc_derivative) * logdif(pr, prob)
         If x > oneMinusP Then
            If oneMinusP - dif < lop Then
               dif = (oneMinusP - lop) * 0.9
            ElseIf oneMinusP - dif > hip Then
               dif = (oneMinusP - hip) * 0.9
            End If
         ElseIf x + dif < lox Then
            dif = (lox - x) * 0.9
         ElseIf x + dif > hix Then
            dif = (hix - x) * 0.9
         End If
         x = x + dif
         oneMinusP = oneMinusP - dif
      End If
   Loop While ((Abs(pr - prob) > prob * 0.00000000000001) And (Abs(dif) > Abs(Min(x, oneMinusP)) * 0.0000000001))
   inv_beta_nc1 = x
End Function

Private Function comp_inv_beta_nc1(ByVal prob As Double, ByVal a As Double, ByVal b As Double, ByVal nc As Double, ByRef oneMinusP As Double) As Double
'Uses approx in A&S 26.6.26 for to get initial estimate the modified NR to improve it.
Dim x As Double, y As Double, pr As Double, dif As Double, temp As Double
Dim hip As Double, lop As Double
Dim hix As Double, lox As Double, nc_derivative As Double
   If (prob > 0.5) Then
      comp_inv_beta_nc1 = inv_beta_nc1(1# - prob, a, b, nc, oneMinusP)
      Exit Function
   End If

   lop = 0#
   hip = 1#
   lox = 0#
   hix = 1#
   pr = Exp(-nc)
   If pr > prob Then
      If 2# * prob > pr Then
         x = invbeta(a + cSmall, b, (pr - prob) / pr, oneMinusP)
      Else
         x = invcompbeta(a + cSmall, b, prob / pr, oneMinusP)
      End If
      If oneMinusP < cSmall Then
         oneMinusP = cSmall
         pr = comp_beta_nc1(x, oneMinusP, a, b, nc, nc_derivative)
         If pr > prob Then
            comp_inv_beta_nc1 = 1#
            oneMinusP = 0#
            Exit Function
         End If
      Else
         temp = oneMinusP
         y = invcompbeta(a + nc ^ 2 / (a + 2 * nc), b, prob, oneMinusP)
         oneMinusP = (a + nc) * oneMinusP / (a + nc * (1# + y))
         If temp < oneMinusP Then
            oneMinusP = temp
         Else
            x = (a + 2# * nc) * y / (a + nc * (1# + y))
         End If
         If oneMinusP < cSmall Then
            oneMinusP = cSmall
            pr = comp_beta_nc1(x, oneMinusP, a, b, nc, nc_derivative)
            If pr > prob Then
               comp_inv_beta_nc1 = 1#
               oneMinusP = 0#
               Exit Function
            End If
         ElseIf x < cSmall Then
            x = cSmall
            pr = comp_beta_nc1(x, oneMinusP, a, b, nc, nc_derivative)
            If pr < prob Then
               comp_inv_beta_nc1 = 0#
               oneMinusP = 1#
               Exit Function
            End If
         End If
      End If
   Else
      y = invcompbeta(a + nc ^ 2 / (a + 2 * nc), b, prob, oneMinusP)
      x = (a + 2# * nc) * y / (a + nc * (1# + y))
      oneMinusP = (a + nc) * oneMinusP / (a + nc * (1# + y))
      If oneMinusP < cSmall Then
         oneMinusP = cSmall
         pr = comp_beta_nc1(x, oneMinusP, a, b, nc, nc_derivative)
         If pr > prob Then
            comp_inv_beta_nc1 = 1#
            oneMinusP = 0#
            Exit Function
         End If
      ElseIf x < cSmall Then
         x = cSmall
         pr = comp_beta_nc1(x, oneMinusP, a, b, nc, nc_derivative)
         If pr < prob Then
            comp_inv_beta_nc1 = 0#
            oneMinusP = 1#
            Exit Function
         End If
      End If
   End If
   dif = x
   Do
      pr = comp_beta_nc1(x, oneMinusP, a, b, nc, nc_derivative)
      If pr < 3E-308 And nc_derivative = 0# Then
         lop = oneMinusP
         hix = x
         dif = dif / 2#
         x = x - dif
         oneMinusP = oneMinusP + dif
      ElseIf nc_derivative = 0# Then
         hip = oneMinusP
         lox = x
         dif = dif / 2#
         x = x - dif
         oneMinusP = oneMinusP + dif
      Else
         If pr < prob Then
            lop = oneMinusP
            hix = x
         Else
            hip = oneMinusP
            lox = x
         End If
         dif = (pr / nc_derivative) * logdif(pr, prob)
         If x > oneMinusP Then
            If oneMinusP - dif < lop Then
               dif = (oneMinusP - lop) * 0.9
            ElseIf oneMinusP - dif > hip Then
               dif = (oneMinusP - hip) * 0.9
            End If
         ElseIf x + dif < lox Then
            dif = (lox - x) * 0.9
         ElseIf x + dif > hix Then
            dif = (hix - x) * 0.9
         End If
         x = x + dif
         oneMinusP = oneMinusP - dif
      End If
   Loop While ((Abs(pr - prob) > prob * 0.00000000000001) And (Abs(dif) > Abs(Min(x, oneMinusP)) * 0.0000000001))
   comp_inv_beta_nc1 = x
End Function

Private Function invBetaLessThanX(ByVal prob As Double, ByVal x As Double, ByVal y As Double, ByVal a As Double, ByVal b As Double) As Double
   Dim oneMinusP As Double
   If x >= y Then
      If invcompbeta(b, a, prob, oneMinusP) >= y * (1# - 0.000000000000001) Then
         invBetaLessThanX = 0#
      Else
         invBetaLessThanX = [#VALUE!]
      End If
   ElseIf invbeta(a, b, prob, oneMinusP) <= x * (1# + 0.000000000000001) Then
      invBetaLessThanX = 0#
   Else
      invBetaLessThanX = [#VALUE!]
   End If
End Function

Private Function compInvBetaLessThanX(ByVal prob As Double, ByVal x As Double, ByVal y As Double, ByVal a As Double, ByVal b As Double) As Double
   Dim oneMinusP As Double
   If x >= y Then
      If invbeta(b, a, prob, oneMinusP) >= y * (1# - 0.000000000000001) Then
         compInvBetaLessThanX = 0#
      Else
         compInvBetaLessThanX = [#VALUE!]
      End If
   ElseIf invcompbeta(a, b, prob, oneMinusP) <= x * (1# + 0.000000000000001) Then
      compInvBetaLessThanX = 0#
   Else
      compInvBetaLessThanX = [#VALUE!]
   End If
End Function

Private Function ncp_beta_nc1(ByVal prob As Double, ByVal x As Double, ByVal y As Double, ByVal a As Double, ByVal b As Double) As Double
'Uses Normal approx for difference of 2 a Negative Binomial and a poisson distributed variable to get initial estimate the modified NR to improve it.
Dim ncp As Double, pr As Double, dif As Double, temp As Double, deriv As Double, c As Double, d As Double, e As Double, sqarg As Double, checked_nc_limit As Boolean, checked_0_limit As Boolean
Dim hi As Double, lo As Double, nc_derivative As Double
   If (prob > 0.5) Then
      ncp_beta_nc1 = comp_ncp_beta_nc1(1# - prob, x, y, a, b)
      Exit Function
   End If

   lo = 0#
   hi = nc_limit
   checked_0_limit = False
   checked_nc_limit = False
   temp = inv_normal(prob) ^ 2
   c = b * x / y
   d = temp - 2# * (a - c)
   If d < 2 * nc_limit Then
      e = (c - a) ^ 2 - temp * c / y
      sqarg = d ^ 2 - 4 * e
      If sqarg < 0 Then
         ncp = d / 2
      Else
         ncp = (d + Sqr(sqarg)) / 2
      End If
   Else
      ncp = nc_limit
   End If
   ncp = Min(Max(0#, ncp), nc_limit)
   If x > y Then
      pr = compbeta(y * (1 + ncp / (ncp + a)) / (1 + ncp / (ncp + a) * y), b, a + ncp ^ 2 / (2 * ncp + a))
   Else
      pr = beta(x / (1 + ncp / (ncp + a) * y), a + ncp ^ 2 / (2 * ncp + a), b)
   End If
   'Debug.Print "ncp_beta_nc1 ncp1 ", ncp, pr
   If ncp = 0# Then
      If pr < prob Then
         ncp_beta_nc1 = invBetaLessThanX(prob, x, y, a, b)
         Exit Function
      Else
         checked_0_limit = True
      End If
   End If
   temp = Min(Max(0#, invcompgamma(b * x, prob) / y - a), nc_limit)
   If temp = ncp Then
      c = pr
   ElseIf x > y Then
      c = compbeta(y * (1 + temp / (temp + a)) / (1 + temp / (temp + a) * y), b, a + temp ^ 2 / (2 * temp + a))
   Else
      c = beta(x / (1 + temp / (temp + a) * y), a + temp ^ 2 / (2 * temp + a), b)
   End If
   'Debug.Print "ncp_beta_nc1 ncp2 ", temp, c
   If temp = 0# Then
      If c < prob Then
         ncp_beta_nc1 = invBetaLessThanX(prob, x, y, a, b)
         Exit Function
      Else
         checked_0_limit = True
      End If
   End If
   If pr * c = 0# Then
      ncp = Min(ncp, temp)
      pr = Max(pr, c)
      If pr = 0# Then
         c = compbeta(y, b, a)
         If c < prob Then
            ncp_beta_nc1 = invBetaLessThanX(prob, x, y, a, b)
            Exit Function
         Else
            checked_0_limit = True
         End If
      End If
   ElseIf Abs(Log(pr / prob)) > Abs(Log(c / prob)) Then
      ncp = temp
      pr = c
   End If
   If ncp = 0# Then
      If b > 1.000001 Then
         deriv = comp_beta_nc1(x, y, a + 1#, b - 1#, ncp, nc_derivative)
         deriv = nc_derivative * y ^ 2 / (b - 1#)
      Else
         deriv = pr - beta_nc1(x, y, a + 1#, b, ncp, nc_derivative)
      End If
      If deriv = 0# Then
         ncp = nc_limit
      Else
         ncp = (pr - prob) / deriv
         If ncp >= nc_limit Then
            ncp = (pr / deriv) * logdif(pr, prob)
         End If
      End If
   Else
      If ncp = nc_limit Then
         If pr > prob Then
            ncp_beta_nc1 = [#VALUE!]
            Exit Function
         Else
            checked_nc_limit = True
         End If
      End If
      If pr > 0 Then
         temp = ncp * 0.999999 'Use numerical derivative on approximation since cheap compared to evaluating non-central beta
         If x > y Then
            c = compbeta(y * (1# + temp / (temp + a)) / (1 + temp / (temp + a) * y), b, a + temp ^ 2 / (2 * temp + a))
         Else
            c = beta(x / (1 + temp / (temp + a) * y), a + temp ^ 2 / (2 * temp + a), b)
         End If
         If pr <> c Then
            dif = (0.000001 * ncp * pr / (pr - c)) * logdif(pr, prob)
            If ncp - dif < 0# Then
               ncp = ncp / 2#
            ElseIf ncp - dif > nc_limit Then
               ncp = (ncp + nc_limit) / 2#
            Else
               ncp = ncp - dif
            End If
         End If
      Else
         ncp = ncp / 2#
      End If
   End If
   dif = ncp
   Do
      pr = beta_nc1(x, y, a, b, ncp, nc_derivative)
      'Debug.Print ncp, pr, prob
      If b > 1.000001 Then
         deriv = beta_nc1(x, y, a + 1#, b - 1#, ncp, nc_derivative)
         deriv = nc_derivative * y ^ 2 / (b - 1#)
      Else
         deriv = pr - beta_nc1(x, y, a + 1#, b, ncp, nc_derivative)
      End If
      If pr < 3E-308 And deriv = 0# Then
         hi = ncp
         dif = dif / 2#
         ncp = ncp - dif
      ElseIf deriv = 0# Then
         lo = ncp
         dif = dif / 2#
         ncp = ncp - dif
      Else
         If pr < prob Then
            hi = ncp
         Else
            lo = ncp
         End If
         dif = (pr / deriv) * logdif(pr, prob)
         If ncp + dif < lo Then
            dif = (lo - ncp) / 2#
            If Not checked_0_limit And (lo = 0#) Then
               temp = cdf_beta_nc(x, a, b, lo)
               If temp < prob Then
                  ncp_beta_nc1 = invBetaLessThanX(prob, x, y, a, b)
                  Exit Function
               Else
                  checked_0_limit = True
               End If
            End If
         ElseIf ncp + dif > hi Then
            dif = (hi - ncp) / 2#
            If Not checked_nc_limit And (hi = nc_limit) Then
               temp = cdf_beta_nc(x, a, b, hi)
               If temp > prob Then
                  ncp_beta_nc1 = [#VALUE!]
                  Exit Function
               Else
                  checked_nc_limit = True
               End If
            End If
         End If
         ncp = ncp + dif
      End If
   Loop While ((Abs(pr - prob) > prob * 0.00000000000001) And (Abs(dif) > Abs(ncp) * 0.0000000001))
   ncp_beta_nc1 = ncp
   'Debug.Print "ncp_beta_nc1", ncp_beta_nc1
End Function

Private Function comp_ncp_beta_nc1(ByVal prob As Double, ByVal x As Double, ByVal y As Double, ByVal a As Double, ByVal b As Double) As Double
'Uses Normal approx for difference of 2 a Negative Binomial and a poisson distributed variable to get initial estimate the modified NR to improve it.
Dim ncp As Double, pr As Double, dif As Double, temp As Double, deriv As Double, c As Double, d As Double, e As Double, sqarg As Double, checked_nc_limit As Boolean, checked_0_limit As Boolean
Dim hi As Double, lo As Double, nc_derivative As Double
   If (prob > 0.5) Then
      comp_ncp_beta_nc1 = ncp_beta_nc1(1# - prob, x, y, a, b)
      Exit Function
   End If

   lo = 0#
   hi = nc_limit
   checked_0_limit = False
   checked_nc_limit = False
   temp = inv_normal(prob) ^ 2
   c = b * x / y
   d = temp - 2# * (a - c)
   If d < 4 * nc_limit Then
      sqarg = d ^ 2 - 4 * e
      If sqarg < 0 Then
         ncp = d / 2
      Else
         ncp = (d - Sqr(sqarg)) / 2
      End If
   Else
      ncp = 0#
   End If
   ncp = Min(Max(0#, ncp), nc_limit)
   If x > y Then
      pr = beta(y * (1 + ncp / (ncp + a)) / (1 + ncp / (ncp + a) * y), b, a + ncp ^ 2 / (2 * ncp + a))
   Else
      pr = compbeta(x / (1 + ncp / (ncp + a) * y), a + ncp ^ 2 / (2 * ncp + a), b)
   End If
   'Debug.Print "comp_ncp_beta_nc1 ncp1 ", ncp, pr
   If ncp = 0# Then
      If pr > prob Then
         comp_ncp_beta_nc1 = compInvBetaLessThanX(prob, x, y, a, b)
         Exit Function
      Else
         checked_0_limit = True
      End If
   End If
   temp = Min(Max(0#, invgamma(b * x, prob) / y - a), nc_limit)
   If temp = ncp Then
      c = pr
   ElseIf x > y Then
      c = beta(y * (1 + temp / (temp + a)) / (1 + temp / (temp + a) * y), b, a + temp ^ 2 / (2 * temp + a))
   Else
      c = compbeta(x / (1 + temp / (temp + a) * y), a + temp ^ 2 / (2 * temp + a), b)
   End If
   'Debug.Print "comp_ncp_beta_nc1 ncp2 ", temp, c
   If temp = 0# Then
      If c > prob Then
         comp_ncp_beta_nc1 = compInvBetaLessThanX(prob, x, y, a, b)
         Exit Function
      Else
         checked_0_limit = True
      End If
   End If
   If pr * c = 0# Then
      ncp = Max(ncp, temp)
      pr = Max(pr, c)
   ElseIf Abs(Log(pr / prob)) > Abs(Log(c / prob)) Then
      ncp = temp
      pr = c
   End If
   If ncp = 0# Then
      If pr > prob Then
         comp_ncp_beta_nc1 = compInvBetaLessThanX(prob, x, y, a, b)
         Exit Function
      Else
         If b > 1.000001 Then
            deriv = beta_nc1(x, y, a + 1#, b - 1#, 0#, nc_derivative)
            deriv = nc_derivative * y ^ 2 / (b - 1#)
         Else
            deriv = comp_beta_nc1(x, y, a + 1#, b, 0#, nc_derivative) - pr
         End If
         If deriv = 0# Then
            ncp = nc_limit
         Else
            ncp = (prob - pr) / deriv
            If ncp >= nc_limit Then
               ncp = -(pr / deriv) * logdif(pr, prob)
            End If
         End If
         checked_0_limit = True
      End If
   Else
      If ncp = nc_limit Then
         If pr < prob Then
            comp_ncp_beta_nc1 = [#VALUE!]
            Exit Function
         Else
            checked_nc_limit = True
         End If
      End If
      If pr > 0 Then
         temp = ncp * 0.999999 'Use numerical derivative on approximation since cheap compared to evaluating non-central beta
         If x > y Then
            c = beta(y * (1# + temp / (temp + a)) / (1 + temp / (temp + a) * y), b, a + temp ^ 2 / (2 * temp + a))
         Else
            c = compbeta(x / (1 + temp / (temp + a) * y), a + temp ^ 2 / (2 * temp + a), b)
         End If
         If pr <> c Then
            dif = -(0.000001 * ncp * pr / (pr - c)) * logdif(pr, prob)
            If ncp + dif < 0 Then
               ncp = ncp / 2
            ElseIf ncp + dif > nc_limit Then
               ncp = (ncp + nc_limit) / 2
            Else
               ncp = ncp + dif
            End If
         End If
      Else
         ncp = (nc_limit + ncp) / 2#
      End If
   End If
   dif = ncp
   Do
      pr = comp_beta_nc1(x, y, a, b, ncp, nc_derivative)
      'Debug.Print ncp, pr, prob
      If b > 1.000001 Then
         deriv = beta_nc1(x, y, a + 1#, b - 1#, ncp, nc_derivative)
         deriv = nc_derivative * y ^ 2 / (b - 1#)
      Else
         deriv = comp_beta_nc1(x, y, a + 1#, b, ncp, nc_derivative) - pr
      End If
      If pr < 3E-308 And deriv = 0# Then
         lo = ncp
         dif = dif / 2#
         ncp = ncp + dif
      ElseIf deriv = 0# Then
         hi = ncp
         dif = dif / 2#
         ncp = ncp - dif
      Else
         If pr < prob Then
            lo = ncp
         Else
            hi = ncp
         End If
         dif = -(pr / deriv) * logdif(pr, prob)
         If ncp + dif < lo Then
            dif = (lo - ncp) / 2#
            If Not checked_0_limit And (lo = 0#) Then
               temp = comp_cdf_beta_nc(x, a, b, lo)
               If temp > prob Then
                  comp_ncp_beta_nc1 = compInvBetaLessThanX(prob, x, y, a, b)
                  Exit Function
               Else
                  checked_0_limit = True
               End If
            End If
         ElseIf ncp + dif > hi Then
            dif = (hi - ncp) / 2#
            If Not checked_nc_limit And (hi = nc_limit) Then
               temp = comp_cdf_beta_nc(x, a, b, hi)
               If temp < prob Then
                  comp_ncp_beta_nc1 = [#VALUE!]
                  Exit Function
               Else
                  checked_nc_limit = True
               End If
            End If
         End If
         ncp = ncp + dif
      End If
   Loop While ((Abs(pr - prob) > prob * 0.00000000000001) And (Abs(dif) > Abs(ncp) * 0.0000000001))
   comp_ncp_beta_nc1 = ncp
   'Debug.Print "comp_ncp_beta_nc1", comp_ncp_beta_nc1
End Function

Public Function pdf_beta_nc(ByVal x As Double, ByVal shape_param1 As Double, ByVal shape_param2 As Double, ByVal nc_param As Double) As Double
  If (shape_param1 < 0#) Or (shape_param2 < 0#) Or (nc_param < 0#) Or (nc_param > nc_limit) Or ((shape_param1 = 0#) And (shape_param2 = 0#)) Then
     pdf_beta_nc = [#VALUE!]
  ElseIf (x < 0# Or x > 1#) Then
     pdf_beta_nc = 0#
  ElseIf (x = 0# Or nc_param = 0#) Then
     pdf_beta_nc = Exp(-nc_param) * pdf_beta(x, shape_param1, shape_param2)
  ElseIf (x = 1# And shape_param2 = 1#) Then
     pdf_beta_nc = shape_param1 + nc_param
  ElseIf (x = 1#) Then
     pdf_beta_nc = pdf_beta(x, shape_param1, shape_param2)
  Else
     Dim nc_derivative As Double
     If (shape_param1 < 1# Or x * shape_param2 <= (1# - x) * (shape_param1 + nc_param)) Then
        pdf_beta_nc = beta_nc1(x, 1# - x, shape_param1, shape_param2, nc_param, nc_derivative)
     Else
        pdf_beta_nc = comp_beta_nc1(x, 1# - x, shape_param1, shape_param2, nc_param, nc_derivative)
     End If
     pdf_beta_nc = nc_derivative
  End If
  pdf_beta_nc = GetRidOfMinusZeroes(pdf_beta_nc)
End Function

Public Function cdf_beta_nc(ByVal x As Double, ByVal shape_param1 As Double, ByVal shape_param2 As Double, ByVal nc_param As Double) As Double
  Dim nc_derivative As Double
  If (shape_param1 < 0#) Or (shape_param2 < 0#) Or (nc_param < 0#) Or (nc_param > nc_limit) Or ((shape_param1 = 0#) And (shape_param2 = 0#)) Then
     cdf_beta_nc = [#VALUE!]
  ElseIf (x < 0#) Then
     cdf_beta_nc = 0#
  ElseIf (x >= 1#) Then
     cdf_beta_nc = 1#
  ElseIf (x = 0# And shape_param1 = 0#) Then
     cdf_beta_nc = Exp(-nc_param)
  ElseIf (x = 0#) Then
     cdf_beta_nc = 0#
  ElseIf (nc_param = 0#) Then
     cdf_beta_nc = beta(x, shape_param1, shape_param2)
  ElseIf (shape_param1 < 1# Or x * shape_param2 <= (1# - x) * (shape_param1 + nc_param)) Then
     cdf_beta_nc = beta_nc1(x, 1# - x, shape_param1, shape_param2, nc_param, nc_derivative)
  Else
     cdf_beta_nc = 1# - comp_beta_nc1(x, 1# - x, shape_param1, shape_param2, nc_param, nc_derivative)
  End If
  cdf_beta_nc = GetRidOfMinusZeroes(cdf_beta_nc)
End Function

Public Function comp_cdf_beta_nc(ByVal x As Double, ByVal shape_param1 As Double, ByVal shape_param2 As Double, ByVal nc_param As Double) As Double
  Dim nc_derivative As Double
  If (shape_param1 < 0#) Or (shape_param2 < 0#) Or (nc_param < 0#) Or (nc_param > nc_limit) Or ((shape_param1 = 0#) And (shape_param2 = 0#)) Then
     comp_cdf_beta_nc = [#VALUE!]
  ElseIf (x < 0#) Then
     comp_cdf_beta_nc = 1#
  ElseIf (x >= 1#) Then
     comp_cdf_beta_nc = 0#
  ElseIf (x = 0# And shape_param1 = 0#) Then
     comp_cdf_beta_nc = -expm1(-nc_param)
  ElseIf (x = 0#) Then
     comp_cdf_beta_nc = 1#
  ElseIf (nc_param = 0#) Then
     comp_cdf_beta_nc = compbeta(x, shape_param1, shape_param2)
  ElseIf (shape_param1 < 1# Or x * shape_param2 >= (1# - x) * (shape_param1 + nc_param)) Then
     comp_cdf_beta_nc = comp_beta_nc1(x, 1# - x, shape_param1, shape_param2, nc_param, nc_derivative)
  Else
     comp_cdf_beta_nc = 1# - beta_nc1(x, 1# - x, shape_param1, shape_param2, nc_param, nc_derivative)
  End If
  comp_cdf_beta_nc = GetRidOfMinusZeroes(comp_cdf_beta_nc)
End Function

Public Function inv_beta_nc(ByVal prob As Double, ByVal shape_param1 As Double, ByVal shape_param2 As Double, ByVal nc_param As Double) As Double
  Dim oneMinusP As Double
  If (shape_param1 < 0#) Or (shape_param2 <= 0#) Or (nc_param < 0#) Or (nc_param > nc_limit) Or (prob < 0#) Or (prob > 1#) Then
     inv_beta_nc = [#VALUE!]
  ElseIf (prob = 0# Or shape_param1 = 0# And prob <= Exp(-nc_param)) Then
     inv_beta_nc = 0#
  ElseIf (prob = 1#) Then
     inv_beta_nc = 1#
  ElseIf (nc_param = 0#) Then
     inv_beta_nc = invbeta(shape_param1, shape_param2, prob, oneMinusP)
  Else
     inv_beta_nc = inv_beta_nc1(prob, shape_param1, shape_param2, nc_param, oneMinusP)
  End If
  inv_beta_nc = GetRidOfMinusZeroes(inv_beta_nc)
End Function

Public Function comp_inv_beta_nc(ByVal prob As Double, ByVal shape_param1 As Double, ByVal shape_param2 As Double, ByVal nc_param As Double) As Double
  Dim oneMinusP As Double
  If (shape_param1 < 0#) Or (shape_param2 <= 0#) Or (nc_param < 0#) Or (nc_param > nc_limit) Or (prob < 0#) Or (prob > 1#) Then
     comp_inv_beta_nc = [#VALUE!]
  ElseIf (prob = 1# Or shape_param1 = 0# And prob >= -expm1(-nc_param)) Then
     comp_inv_beta_nc = 0#
  ElseIf (prob = 0#) Then
     comp_inv_beta_nc = 1#
  ElseIf (nc_param = 0#) Then
     comp_inv_beta_nc = invcompbeta(shape_param1, shape_param2, prob, oneMinusP)
  Else
     comp_inv_beta_nc = comp_inv_beta_nc1(prob, shape_param1, shape_param2, nc_param, oneMinusP)
  End If
  comp_inv_beta_nc = GetRidOfMinusZeroes(comp_inv_beta_nc)
End Function

Public Function ncp_beta_nc(ByVal prob As Double, ByVal x As Double, ByVal shape_param1 As Double, ByVal shape_param2 As Double) As Double
  If (shape_param1 < 0#) Or (shape_param2 <= 0#) Or (x < 0#) Or (x >= 1#) Or (prob <= 0#) Or (prob > 1#) Then
     ncp_beta_nc = [#VALUE!]
  ElseIf (x = 0# And shape_param1 = 0#) Then
     ncp_beta_nc = -Log(prob)
  ElseIf (x = 0# Or prob = 1#) Then
     ncp_beta_nc = [#VALUE!]
  Else
     ncp_beta_nc = ncp_beta_nc1(prob, x, 1# - x, shape_param1, shape_param2)
  End If
  ncp_beta_nc = GetRidOfMinusZeroes(ncp_beta_nc)
End Function

Public Function comp_ncp_beta_nc(ByVal prob As Double, ByVal x As Double, ByVal shape_param1 As Double, ByVal shape_param2 As Double) As Double
  If (shape_param1 < 0#) Or (shape_param2 <= 0#) Or (x < 0#) Or (x >= 1#) Or (prob < 0#) Or (prob >= 1#) Then
     comp_ncp_beta_nc = [#VALUE!]
  ElseIf (x = 0# And shape_param1 = 0#) Then
     comp_ncp_beta_nc = -log0(-prob)
  ElseIf (x = 0# Or prob = 0#) Then
     comp_ncp_beta_nc = [#VALUE!]
  Else
     comp_ncp_beta_nc = comp_ncp_beta_nc1(prob, x, 1# - x, shape_param1, shape_param2)
  End If
  comp_ncp_beta_nc = GetRidOfMinusZeroes(comp_ncp_beta_nc)
End Function

Public Function pdf_fdist_nc(ByVal x As Double, ByVal df1 As Double, ByVal df2 As Double, ByVal nc As Double) As Double
   df1 = AlterForIntegralChecks_df(df1)
   df2 = AlterForIntegralChecks_df(df2)
   If (df1 <= 0# Or df2 <= 0# Or (nc < 0#) Or (nc > 2# * nc_limit)) Then
      pdf_fdist_nc = [#VALUE!]
   ElseIf (x < 0#) Then
      pdf_fdist_nc = 0#
   ElseIf (x = 0# Or nc = 0#) Then
      pdf_fdist_nc = Exp(-nc / 2#) * pdf_fdist(x, df1, df2)
   Else
      Dim p As Double, q As Double, nc_derivative As Double
      If x > 1# Then
         q = df2 / x
         p = q + df1
         q = q / p
         p = df1 / p
      Else
         p = df1 * x
         q = df2 + p
         p = p / q
         q = df2 / q
      End If
      If (df1 < 1# Or p * df2 <= q * (df1 + nc)) Then
         pdf_fdist_nc = beta_nc1(p, q, df1 / 2#, df2 / 2#, nc / 2#, nc_derivative)
      Else
         pdf_fdist_nc = comp_beta_nc1(p, q, df1 / 2#, df2 / 2#, nc / 2#, nc_derivative)
      End If
      pdf_fdist_nc = (nc_derivative * q) * (df1 * q / df2)
   End If
   pdf_fdist_nc = GetRidOfMinusZeroes(pdf_fdist_nc)
End Function

Public Function cdf_fdist_nc(ByVal x As Double, ByVal df1 As Double, ByVal df2 As Double, ByVal nc As Double) As Double
   df1 = AlterForIntegralChecks_df(df1)
   df2 = AlterForIntegralChecks_df(df2)
   If (df1 <= 0# Or df2 <= 0# Or (nc < 0#) Or (nc > 2# * nc_limit)) Then
      cdf_fdist_nc = [#VALUE!]
   ElseIf (x <= 0#) Then
      cdf_fdist_nc = 0#
   Else
      Dim p As Double, q As Double, nc_derivative As Double
      If x > 1# Then
         q = df2 / x
         p = q + df1
         q = q / p
         p = df1 / p
      Else
         p = df1 * x
         q = df2 + p
         p = p / q
         q = df2 / q
      End If
      'If p < cSmall And x <> 0# Or q < cSmall Then
      '   cdf_fdist_nc = [#VALUE!]
      '   Exit Function
      'End If
      df2 = df2 / 2#
      df1 = df1 / 2#
      nc = nc / 2#
      If (nc = 0# And p <= q) Then
         cdf_fdist_nc = beta(p, df1, df2)
      ElseIf (nc = 0#) Then
         cdf_fdist_nc = compbeta(q, df2, df1)
      ElseIf (df1 < 1# Or p * df2 <= q * (df1 + nc)) Then
         cdf_fdist_nc = beta_nc1(p, q, df1, df2, nc, nc_derivative)
      Else
         cdf_fdist_nc = 1# - comp_beta_nc1(p, q, df1, df2, nc, nc_derivative)
      End If
   End If
   cdf_fdist_nc = GetRidOfMinusZeroes(cdf_fdist_nc)
End Function

Public Function comp_cdf_fdist_nc(ByVal x As Double, ByVal df1 As Double, ByVal df2 As Double, ByVal nc As Double) As Double
   df1 = AlterForIntegralChecks_df(df1)
   df2 = AlterForIntegralChecks_df(df2)
   If (df1 <= 0# Or df2 <= 0# Or (nc < 0#) Or (nc > 2# * nc_limit)) Then
      comp_cdf_fdist_nc = [#VALUE!]
   ElseIf (x <= 0#) Then
      comp_cdf_fdist_nc = 1#
   Else
      Dim p As Double, q As Double, nc_derivative As Double
      If x > 1# Then
         q = df2 / x
         p = q + df1
         q = q / p
         p = df1 / p
      Else
         p = df1 * x
         q = df2 + p
         p = p / q
         q = df2 / q
      End If
      'If p < cSmall And x <> 0# Or q < cSmall Then
      '   comp_cdf_fdist_nc = [#VALUE!]
      '   Exit Function
      'End If
      df2 = df2 / 2#
      df1 = df1 / 2#
      nc = nc / 2#
      If (nc = 0# And p <= q) Then
         comp_cdf_fdist_nc = compbeta(p, df1, df2)
      ElseIf (nc = 0#) Then
         comp_cdf_fdist_nc = beta(q, df2, df1)
      ElseIf (df1 < 1# Or p * df2 >= q * (df1 + nc)) Then
         comp_cdf_fdist_nc = comp_beta_nc1(p, q, df1, df2, nc, nc_derivative)
      Else
         comp_cdf_fdist_nc = 1# - beta_nc1(p, q, df1, df2, nc, nc_derivative)
      End If
   End If
   comp_cdf_fdist_nc = GetRidOfMinusZeroes(comp_cdf_fdist_nc)
End Function

Public Function inv_fdist_nc(ByVal prob As Double, ByVal df1 As Double, ByVal df2 As Double, ByVal nc As Double) As Double
   df1 = AlterForIntegralChecks_df(df1)
   df2 = AlterForIntegralChecks_df(df2)
   If (df1 <= 0# Or df2 <= 0# Or (nc < 0#) Or (nc > 2# * nc_limit) Or prob < 0# Or prob >= 1#) Then
      inv_fdist_nc = [#VALUE!]
   ElseIf (prob = 0#) Then
      inv_fdist_nc = 0#
   Else
      Dim temp As Double, oneMinusP As Double
      df1 = df1 / 2#
      df2 = df2 / 2#
      If nc = 0# Then
         temp = invbeta(df1, df2, prob, oneMinusP)
      Else
         temp = inv_beta_nc1(prob, df1, df2, nc / 2#, oneMinusP)
      End If
      inv_fdist_nc = df2 * temp / (df1 * oneMinusP)
      'If oneMinusP < cSmall Then inv_fdist_nc = [#VALUE!]
   End If
   inv_fdist_nc = GetRidOfMinusZeroes(inv_fdist_nc)
End Function

Public Function comp_inv_fdist_nc(ByVal prob As Double, ByVal df1 As Double, ByVal df2 As Double, ByVal nc As Double) As Double
   df1 = AlterForIntegralChecks_df(df1)
   df2 = AlterForIntegralChecks_df(df2)
   If (df1 <= 0# Or df2 <= 0# Or (nc < 0#) Or (nc > 2# * nc_limit) Or prob <= 0# Or prob > 1#) Then
      comp_inv_fdist_nc = [#VALUE!]
   ElseIf (prob = 1#) Then
      comp_inv_fdist_nc = 0#
   Else
      Dim temp As Double, oneMinusP As Double
      df1 = df1 / 2#
      df2 = df2 / 2#
      If nc = 0# Then
         temp = invcompbeta(df1, df2, prob, oneMinusP)
      Else
         temp = comp_inv_beta_nc1(prob, df1, df2, nc / 2#, oneMinusP)
      End If
      comp_inv_fdist_nc = df2 * temp / (df1 * oneMinusP)
      'If oneMinusP < cSmall Then comp_inv_fdist_nc = [#VALUE!]
   End If
   comp_inv_fdist_nc = GetRidOfMinusZeroes(comp_inv_fdist_nc)
End Function

Public Function ncp_fdist_nc(ByVal prob As Double, ByVal x As Double, ByVal df1 As Double, ByVal df2 As Double) As Double
  df1 = AlterForIntegralChecks_df(df1)
  df2 = AlterForIntegralChecks_df(df2)
  If (df1 <= 0#) Or (df2 <= 0#) Or (x <= 0#) Or (prob <= 0#) Or (prob >= 1#) Then
     ncp_fdist_nc = [#VALUE!]
  Else
     Dim p As Double, q As Double
     If x > 1# Then
        q = df2 / x
        p = q + df1
        q = q / p
        p = df1 / p
     Else
        p = df1 * x
        q = df2 + p
        p = p / q
        q = df2 / q
     End If
     df2 = df2 / 2#
     df1 = df1 / 2#
     ncp_fdist_nc = ncp_beta_nc1(prob, p, q, df1, df2) * 2#
  End If
  ncp_fdist_nc = GetRidOfMinusZeroes(ncp_fdist_nc)
End Function

Public Function comp_ncp_fdist_nc(ByVal prob As Double, ByVal x As Double, ByVal df1 As Double, ByVal df2 As Double) As Double
  df1 = AlterForIntegralChecks_df(df1)
  df2 = AlterForIntegralChecks_df(df2)
  If (df1 <= 0#) Or (df2 <= 0#) Or (x <= 0#) Or (prob <= 0#) Or (prob >= 1#) Then
     comp_ncp_fdist_nc = [#VALUE!]
  Else
     Dim p As Double, q As Double
     If x > 1# Then
        q = df2 / x
        p = q + df1
        q = q / p
        p = df1 / p
     Else
        p = df1 * x
        q = df2 + p
        p = p / q
        q = df2 / q
     End If
     df1 = df1 / 2#
     df2 = df2 / 2#
     comp_ncp_fdist_nc = comp_ncp_beta_nc1(prob, p, q, df1, df2) * 2#
  End If
  comp_ncp_fdist_nc = GetRidOfMinusZeroes(comp_ncp_fdist_nc)
End Function

Private Function t_nc1(ByVal t As Double, ByVal df As Double, ByVal nct As Double, ByRef nc_derivative As Double) As Double
'y is 1-x but held accurately to avoid possible cancellation errors
'nc_derivative holds t * derivative
   Dim aa As Double, bb As Double, nc_dtemp As Double
   Dim n As Double, p As Double, q As Double, w As Double, V As Double, r As Double, s As Double, ps As Double
   Dim result1 As Double, result2 As Double, term1 As Double, term2 As Double, ptnc As Double, qtnc As Double, ptx As Double, qtx As Double
   Dim a As Double, b As Double, x As Double, y As Double, nc As Double
   Dim save_result1 As Double, save_result2 As Double, phi As Double, vScale As Double
   phi = cnormal(-Abs(nct))
   a = 0.5
   b = df / 2#
   If Abs(t) >= Min(1#, df) Then
      y = df / t
      x = t + y
      y = y / x
      x = t / x
   Else
      x = t * t
      y = df + x
      x = x / y
      y = df / y
   End If
   If y < cSmall Then
      t_nc1 = [#VALUE!]
      Exit Function
   End If
   nc = nct * nct / 2#
   aa = a - nc * x * (a + b)
   bb = (x * nc - 1#) - a
   If (bb < 0#) Then
      n = bb - Sqr(bb ^ 2 - 4# * aa)
      n = Int(2# * aa / n)
   Else
      n = Int((bb + Sqr(bb ^ 2 - 4# * aa)) / 2#)
   End If
   If n < 0# Then
      n = 0#
   End If
   aa = n + a
   bb = n + 0.5
   qtnc = poissonTerm(bb, nc, nc - bb, 0#)
   bb = n
   ptnc = poissonTerm(bb, nc, nc - bb, 0#)
   ptx = binomialTerm(aa, b, x, y, b * x - aa * y, 0#) / (aa + b) '(I(x, aa, b) - I(x, aa+1, b))/b
   qtx = binomialTerm(aa + 0.5, b, x, y, b * x - (aa + 0.5) * y, 0#) / (aa + b + 0.5) '(I(x, aa+1/2, b) - I(x, aa+3/2, b))/b
   If b > 1# Then
      ptx = b * ptx
      qtx = b * qtx
   End If
   vScale = Max(ptx, qtx)
   If ptx = vScale Then
      s = 1#
   Else
      s = ptx / vScale
   End If
   If qtx = vScale Then
      r = 1#
   Else
      r = qtx / vScale
   End If
   s = (aa + b) * s
   r = (aa + b + 0.5) * r
   aa = aa + 1#
   bb = bb + 1#
   p = nc / bb * ptnc
   q = nc / (bb + 0.5) * qtnc
   ps = p * s + q * r
   nc_derivative = ps
   s = x / aa * s  ' I(x, aa, b) - I(x, aa+1, b)
   r = x / (aa + 0.5) * r ' I(x, aa+1/2, b) - I(x, aa+3/2, b)
   w = p
   V = q
   term1 = s * w
   term2 = r * V
   result1 = term1
   result2 = term2
   While ((((term1 + term2) > 0.000000000000001 * (result1 + result2)) And (p > 1E-16 * w)) Or (ps > 1E-16 * nc_derivative))
       s = (aa + b) * s
       r = (aa + b + 0.5) * r
       aa = aa + 1#
       bb = bb + 1#
       p = nc / bb * p
       q = nc / (bb + 0.5) * q
       ps = p * s + q * r
       nc_derivative = nc_derivative + ps
       s = x / aa * s ' I(x, aa, b) - I(x, aa+1, b)
       r = x / (aa + 0.5) * r ' I(x, aa+1/2, b) - I(x, aa+3/2, b)
       w = w + p
       V = V + q
       term1 = s * w
       term2 = r * V
       result1 = result1 + term1
       result2 = result2 + term2
   Wend
   If x > y Then
      s = compbeta(y, b, a + (bb + 1#))
      r = compbeta(y, b, a + (bb + 1.5))
   Else
      s = beta(x, a + (bb + 1#), b)
      r = beta(x, a + (bb + 1.5), b)
   End If
   nc_derivative = x * nc_derivative * vScale
   If b <= 1# Then vScale = vScale * b
   save_result1 = result1 * vScale + s * w
   save_result2 = result2 * vScale + r * V

   ps = 1#
   nc_dtemp = 0#
   aa = n + a
   bb = n
   vScale = Max(ptnc, qtnc)
   If ptnc = vScale Then
      p = 1#
   Else
      p = ptnc / vScale
   End If
   If qtnc = vScale Then
      q = 1#
   Else
      q = qtnc / vScale
   End If
   s = ptx ' I(x, aa, b) - I(x, aa+1, b)
   r = qtx ' I(x, aa+1/2, b) - I(x, aa+3/2, b)
   If x > y Then
      w = compbeta(y, b, aa) ' I(x, aa, b)
      V = compbeta(y, b, aa + 0.5) ' I(x, aa+1/2, b)
   Else
      w = beta(x, aa, b) ' I(x, aa, b)
      V = beta(x, aa + 0.5, b) ' I(x, aa+1/2, b)
   End If
   term1 = p * w
   term2 = q * V
   result1 = term1
   result2 = term2
   While bb > 0# And ((((term1 + term2) > 0.000000000000001 * (result1 + result2)) And (s > 1E-16 * w)) Or (ps > 1E-16 * nc_dtemp))
       s = aa / x * s
       r = (aa + 0.5) / x * r
       ps = p * s + q * r
       nc_dtemp = nc_dtemp + ps
       p = bb / nc * p
       q = (bb + 0.5) / nc * q
       aa = aa - 1#
       bb = bb - 1#
       If bb = 0# Then aa = a
       s = s / (aa + b) ' I(x, aa, b) - I(x, aa+1, b)
       r = r / (aa + b + 0.5) ' I(x, aa+1/2, b) - I(x, aa+3/2, b)
       If b > 1# Then
          w = w + s ' I(x, aa, b)
          V = V + r ' I(x, aa+0.5, b)
       Else
          w = w + b * s
          V = V + b * r
       End If
       term1 = p * w
       term2 = q * V
       result1 = result1 + term1
       result2 = result2 + term2
   Wend
   nc_dtemp = x * nc_dtemp + p * aa * s + q * (aa + 0.5) * r
   p = cpoisson(bb - 1#, nc, nc - bb + 1#)
   q = cpoisson(bb - 0.5, nc, nc - bb + 0.5) - 2# * phi
   result1 = save_result1 + result1 * vScale + p * w
   result2 = save_result2 + result2 * vScale + q * V
   If t > 0# Then
      t_nc1 = phi + 0.5 * (result1 + result2)
      nc_derivative = nc_derivative + nc_dtemp * vScale
   Else
      t_nc1 = phi - 0.5 * (result1 - result2)
   End If
End Function

Private Function comp_t_nc1(ByVal t As Double, ByVal df As Double, ByVal nct As Double, ByRef nc_derivative As Double) As Double
'y is 1-x but held accurately to avoid possible cancellation errors
'nc_derivative holds t * derivative
   Dim aa As Double, bb As Double, nc_dtemp As Double
   Dim n As Double, p As Double, q As Double, w As Double, V As Double, r As Double, s As Double, ps As Double
   Dim result1 As Double, result2 As Double, term1 As Double, term2 As Double, ptnc As Double, qtnc As Double, ptx As Double, qtx As Double
   Dim a As Double, b As Double, x As Double, y As Double, nc As Double
   Dim save_result1 As Double, save_result2 As Double, vScale As Double
   a = 0.5
   b = df / 2#
   If Abs(t) >= Min(1#, df) Then
      y = df / t
      x = t + y
      y = y / x
      x = t / x
   Else
      x = t * t
      y = df + x
      x = x / y
      y = df / y
   End If
   If y < cSmall Then
      comp_t_nc1 = [#VALUE!]
      Exit Function
   End If
   nc = nct * nct / 2#
   aa = a - nc * x * (a + b)
   bb = (x * nc - 1#) - a
   If (bb < 0#) Then
      n = bb - Sqr(bb ^ 2 - 4# * aa)
      n = Int(2# * aa / n)
   Else
      n = Int((bb + Sqr(bb ^ 2 - 4# * aa)) / 2)
   End If
   If n < 0# Then
      n = 0#
   End If
   aa = n + a
   bb = n + 0.5
   qtnc = poissonTerm(bb, nc, nc - bb, 0#)
   bb = n
   ptnc = poissonTerm(bb, nc, nc - bb, 0#)
   ptx = binomialTerm(aa, b, x, y, b * x - aa * y, 0#) / (aa + b) '((1 - I(x, aa+1, b)) - (1 - I(x, aa, b)))/b
   qtx = binomialTerm(aa + 0.5, b, x, y, b * x - (aa + 0.5) * y, 0#) / (aa + b + 0.5) '((1 - I(x, aa+3/2, b)) - (1 - I(x, aa+1/2, b)))/b
   If b > 1# Then
      ptx = b * ptx
      qtx = b * qtx
   End If
   vScale = Max(ptnc, qtnc)
   If ptnc = vScale Then
      p = 1#
   Else
      p = ptnc / vScale
   End If
   If qtnc = vScale Then
      q = 1#
   Else
      q = qtnc / vScale
   End If
   nc_derivative = 0#
   s = ptx
   r = qtx
   If x > y Then
      V = beta(y, b, aa + 0.5) '  1 - I(x, aa+1/2, b)
      w = beta(y, b, aa) '  1 - I(x, aa, b)
   Else
      V = compbeta(x, aa + 0.5, b) ' 1 - I(x, aa+1/2, b)
      w = compbeta(x, aa, b) ' 1 - I(x, aa, b)
   End If
   term1 = 0#
   term2 = 0#
   result1 = term1
   result2 = term2
   Do
       If b > 1# Then
          w = w + s ' 1 - I(x, aa, b)
          V = V + r ' 1 - I(x, aa+1/2, b)
       Else
          w = w + b * s
          V = V + b * r
       End If
       s = (aa + b) * s
       r = (aa + b + 0.5) * r
       aa = aa + 1#
       bb = bb + 1#
       p = nc / bb * p
       q = nc / (bb + 0.5) * q
       ps = p * s + q * r
       nc_derivative = nc_derivative + ps
       s = x / aa * s ' (1 - I(x, aa+1, b)) - (1 - I(x, aa, b))
       r = x / (aa + 0.5) * r ' (1 - I(x, aa+3/2, b)) - (1 - I(x, aa+1/2, b))
       term1 = p * w
       term2 = q * V
       result1 = result1 + term1
       result2 = result2 + term2
   Loop While ((((term1 + term2) > 0.000000000000001 * (result1 + result2)) And (s > 1E-16 * w)) Or (ps > 1E-16 * nc_derivative))
   p = comppoisson(bb, nc, nc - bb)
   bb = bb + 0.5
   q = comppoisson(bb, nc, nc - bb)
   nc_derivative = x * nc_derivative * vScale
   save_result1 = result1 * vScale + p * w
   save_result2 = result2 * vScale + q * V
   ps = 1#
   nc_dtemp = 0#
   aa = n + a
   bb = n
   p = ptnc
   q = qtnc
   vScale = Max(ptx, qtx)
   If ptx = vScale Then
      s = 1#
   Else
      s = ptx / vScale
   End If
   If qtx = vScale Then
      r = 1#
   Else
      r = qtx / vScale
   End If
   w = p
   V = q
   term1 = 1#
   term2 = 1#
   result1 = 0#
   result2 = 0#
   While bb > 0# And ((((term1 + term2) > 0.000000000000001 * (result1 + result2)) And (p > 1E-16 * w)) Or (ps > 1E-16 * nc_dtemp))
      r = (aa + 0.5) / x * r
      s = aa / x * s
      ps = p * s + q * r
      nc_dtemp = nc_dtemp + ps
      p = bb / nc * p
      q = (bb + 0.5) / nc * q
      aa = aa - 1#
      bb = bb - 1#
      If bb = 0# Then aa = a
      r = r / (aa + b + 0.5) ' (1 - I(x, aa+3/2, b)) - (1 - I(x, aa+1/2, b))
      s = s / (aa + b) ' (1 - I(x, aa + 1, b)) - (1 - I(x, aa, b))
      term1 = s * w
      term2 = r * V
      result1 = result1 + term1
      result2 = result2 + term2
      w = w + p
      V = V + q
   Wend
   nc_dtemp = (x * nc_dtemp + p * aa * s + q * (aa + 0.5) * r) * vScale
   If x > y Then
      r = beta(y, b, a + (bb + 0.5))
      s = beta(y, b, a + bb)
   Else
      r = compbeta(x, a + (bb + 0.5), b)
      s = compbeta(x, a + bb, b)
   End If
   If b <= 1# Then vScale = vScale * b
   result1 = save_result1 + result1 * vScale + s * w
   result2 = save_result2 + result2 * vScale + r * V
   If t > 0# Then
      comp_t_nc1 = 0.5 * (result1 + result2)
      nc_derivative = nc_derivative + nc_dtemp
   Else
      comp_t_nc1 = 1# - 0.5 * (result1 - result2)
   End If
End Function

Private Function inv_t_nc1(ByVal prob As Double, ByVal df As Double, ByVal nc As Double, ByRef oneMinusP As Double) As Double
'Uses approximations in A&S 26.6.26 and 26.7.10 for to get initial estimate, the modified NR to improve it.
Dim x As Double, y As Double, pr As Double, dif As Double, temp As Double, nc_beta_param As Double
Dim hix As Double, lox As Double, test As Double, nc_derivative As Double
   If (prob > 0.5) Then
      inv_t_nc1 = comp_inv_t_nc1(1# - prob, df, nc, oneMinusP)
      Exit Function
   End If
   nc_beta_param = nc ^ 2 / 2#
   lox = 0#
   hix = t_nc_limit * Sqr(df)
   pr = Exp(-nc_beta_param)
   If pr > prob Then
      If 2# * prob > pr Then
         x = invcompbeta(0.5, df / 2#, (pr - prob) / pr, oneMinusP)
      Else
         x = invbeta(0.5, df / 2#, prob / pr, oneMinusP)
      End If
      If x = 0# Then
         inv_t_nc1 = 0#
         Exit Function
      Else
         temp = oneMinusP
         y = invbeta((0.5 + nc_beta_param) ^ 2 / (0.5 + 2# * nc_beta_param), df / 2#, prob, oneMinusP)
         oneMinusP = (0.5 + nc_beta_param) * oneMinusP / (0.5 + nc_beta_param * (1# + y))
         If temp > oneMinusP Then
            oneMinusP = temp
         Else
            x = (0.5 + 2# * nc_beta_param) * y / (0.5 + nc_beta_param * (1# + y))
         End If
         If oneMinusP < cSmall Then
            pr = t_nc1(hix, df, nc, nc_derivative)
            If pr < prob Then
               inv_t_nc1 = [#VALUE!]
               oneMinusP = 0#
               Exit Function
            End If
            oneMinusP = 4# * cSmall
         End If
      End If
   Else
      y = invbeta((0.5 + nc_beta_param) ^ 2 / (0.5 + 2# * nc_beta_param), df / 2#, prob, oneMinusP)
      x = (0.5 + 2# * nc_beta_param) * y / (0.5 + nc_beta_param * (1 + y))
      oneMinusP = (0.5 + nc_beta_param) * oneMinusP / (0.5 + nc_beta_param * (1# + y))
      If oneMinusP < cSmall Then
         pr = t_nc1(hix, df, nc, nc_derivative)
         If pr < prob Then
            inv_t_nc1 = [#VALUE!]
            oneMinusP = 0#
            Exit Function
         End If
         oneMinusP = 4# * cSmall
      End If
   End If
   test = Sqr(df * x) / Sqr(oneMinusP)
   Do
      pr = t_nc1(test, df, nc, nc_derivative)
      If pr < prob Then
         lox = test
      Else
         hix = test
      End If
      If nc_derivative = 0# Then
         If pr < prob Then
            dif = (hix - lox) / 2#
         Else
            dif = (lox - hix) / 2#
         End If
      Else
         dif = -(pr * test / nc_derivative) * logdif(pr, prob)
         If df < 2# Then dif = 2# * dif / df
         If test + dif < lox Then
            If lox = 0 Then
               dif = (lox - test) * 0.9999999999
            Else
               dif = (lox - test) * 0.9
            End If
         ElseIf test + dif > hix Then
            dif = (hix - test) * 0.9
         End If
      End If
      test = test + dif
   Loop While ((Abs(pr - prob) > prob * 0.00000000000001) And (Abs(dif) > test * 0.0000000001))
   inv_t_nc1 = test
End Function

Private Function comp_inv_t_nc1(ByVal prob As Double, ByVal df As Double, ByVal nc As Double, ByRef oneMinusP As Double) As Double
'Uses approximations in A&S 26.6.26 and 26.7.10 for to get initial estimate, the modified NR to improve it.
Dim x As Double, y As Double, pr As Double, dif As Double, temp As Double, nc_beta_param As Double
Dim hix As Double, lox As Double, test As Double, nc_derivative As Double
   If (prob > 0.5) Then
      comp_inv_t_nc1 = inv_t_nc1(1# - prob, df, nc, oneMinusP)
      Exit Function
   End If
   nc_beta_param = nc ^ 2 / 2#
   lox = 0#
   hix = t_nc_limit * Sqr(df)
   pr = Exp(-nc_beta_param)
   If pr > prob Then
      If 2# * prob > pr Then
         x = invbeta(0.5, df / 2#, (pr - prob) / pr, oneMinusP)
      Else
         x = invcompbeta(0.5, df / 2#, prob / pr, oneMinusP)
      End If
      If oneMinusP < cSmall Then
         pr = comp_t_nc1(hix, df, nc, nc_derivative)
         If pr > prob Then
            comp_inv_t_nc1 = [#VALUE!]
            oneMinusP = 0#
            Exit Function
         End If
         oneMinusP = 4# * cSmall
      Else
         temp = oneMinusP
         y = invcompbeta((0.5 + nc_beta_param) ^ 2 / (0.5 + 2# * nc_beta_param), df / 2#, prob, oneMinusP)
         oneMinusP = (0.5 + nc_beta_param) * oneMinusP / (0.5 + nc_beta_param * (1# + y))
         If temp < oneMinusP Then
            oneMinusP = temp
         Else
            x = (0.5 + 2# * nc_beta_param) * y / (0.5 + nc_beta_param * (1# + y))
         End If
         If oneMinusP < cSmall Then
            pr = comp_t_nc1(hix, df, nc, nc_derivative)
            If pr > prob Then
               comp_inv_t_nc1 = [#VALUE!]
               oneMinusP = 0#
               Exit Function
            End If
            oneMinusP = 4# * cSmall
         End If
      End If
   Else
      y = invcompbeta((0.5 + nc_beta_param) ^ 2 / (0.5 + 2# * nc_beta_param), df / 2#, prob, oneMinusP)
      x = (0.5 + 2# * nc_beta_param) * y / (0.5 + nc_beta_param * (1# + y))
      oneMinusP = (0.5 + nc_beta_param) * oneMinusP / (0.5 + nc_beta_param * (1# + y))
      If oneMinusP < cSmall Then
         pr = comp_t_nc1(hix, df, nc, nc_derivative)
         If pr > prob Then
            comp_inv_t_nc1 = [#VALUE!]
            oneMinusP = 0#
            Exit Function
         End If
         oneMinusP = 4# * cSmall
      End If
   End If
   test = Sqr(df * x) / Sqr(oneMinusP)
   dif = test
   Do
      pr = comp_t_nc1(test, df, nc, nc_derivative)
      If pr < prob Then
         hix = test
      Else
         lox = test
      End If
      If nc_derivative = 0# Then
         If pr < prob Then
            dif = (lox - hix) / 2#
         Else
            dif = (hix - lox) / 2#
         End If
      Else
         dif = (pr * test / nc_derivative) * logdif(pr, prob)
         If df < 2# Then dif = 2# * dif / df
         If test + dif < lox Then
            dif = (lox - test) * 0.9
         ElseIf test + dif > hix Then
            dif = (hix - test) * 0.9
         End If
      End If
      test = test + dif
   Loop While ((Abs(pr - prob) > prob * 0.00000000000001) And (Abs(dif) > test * 0.0000000001))
   comp_inv_t_nc1 = test
End Function

Private Function ncp_t_nc1(ByVal prob As Double, ByVal t As Double, ByVal df As Double) As Double
'Uses Normal approx for non-central t (A&S 26.7.10) to get initial estimate the modified NR to improve it.
Dim ncp As Double, pr As Double, dif As Double, temp As Double, deriv As Double, checked_tnc_limit As Boolean, checked_0_limit As Boolean
Dim hi As Double, lo As Double, tnc_limit As Double, x As Double, y As Double
   If (prob > 0.5) Then
      ncp_t_nc1 = comp_ncp_t_nc1(1# - prob, t, df)
      Exit Function
   End If

   lo = 0#
   tnc_limit = Sqr(2# * nc_limit)
   hi = tnc_limit
   checked_0_limit = False
   checked_tnc_limit = False
   If t >= Min(1#, df) Then
      y = df / t
      x = t + y
      y = y / x
      x = t / x
   Else
      x = t * t
      y = df + x
      x = x / y
      y = df / y
   End If
   temp = -inv_normal(prob)
   If t > df Then
        ncp = t * (1# - 0.25 / df) + temp * Sqr(t) * Sqr((1# / t + 0.5 * t / df))
   Else
        ncp = t * (1# - 0.25 / df) + temp * Sqr((1# + (0.5 * t / df) * t))
   End If
   ncp = Max(temp, ncp)
   'Debug.Print "ncp_estimate1", ncp
   If x > 1E-200 Then 'I think we can put more accurate bounds on when this will not deliver a sensible answer
      temp = invcompgamma(0.5 * x * df, prob) / y - 0.5
      If temp > 0 Then
         temp = Sqr(2# * temp)
         If temp > ncp Then
            ncp = temp
         End If
      End If
   End If
   'Debug.Print "ncp_estimate2", ncp
   ncp = Min(ncp, tnc_limit)
   If ncp = tnc_limit Then
      pr = cdf_t_nc(t, df, ncp)
      If pr > prob Then
         ncp_t_nc1 = [#VALUE!]
         Exit Function
      Else
         checked_tnc_limit = True
      End If
   End If
   dif = ncp
   Do
      pr = cdf_t_nc(t, df, ncp)
      'Debug.Print ncp, pr, prob
      If ncp > 1 Then
         deriv = cdf_t_nc(t, df, ncp * (1 - 0.000001))
         deriv = 1000000# * (deriv - pr) / ncp
      ElseIf ncp > 0.000001 Then
         deriv = cdf_t_nc(t, df, ncp + 0.000001)
         deriv = 1000000# * (pr - deriv)
      ElseIf x < y Then
         deriv = comp_cdf_beta(x, 1, df / 2) * OneOverSqrTwoPi
      Else
         deriv = cdf_beta(y, df / 2, 1) * OneOverSqrTwoPi
      End If
      If pr < 3E-308 And deriv = 0# Then
         hi = ncp
         dif = dif / 2#
         ncp = ncp - dif
      ElseIf deriv = 0# Then
         lo = ncp
         dif = dif / 2#
         ncp = ncp - dif
      Else
         If pr < prob Then
            hi = ncp
         Else
            lo = ncp
         End If
         dif = (pr / deriv) * logdif(pr, prob)
         If ncp + dif < lo Then
            dif = (lo - ncp) / 2#
            If Not checked_0_limit And (lo = 0#) Then
               temp = cdf_t_nc(t, df, lo)
               If temp < prob Then
                  If invtdist(prob, df) <= t Then
                     ncp_t_nc1 = 0#
                  Else
                     ncp_t_nc1 = [#VALUE!]
                  End If
                  Exit Function
               Else
                  checked_0_limit = True
               End If
               dif = dif * 1.99999999
            End If
         ElseIf ncp + dif > hi Then
            dif = (hi - ncp) / 2#
            If Not checked_tnc_limit And (hi = tnc_limit) Then
               temp = cdf_t_nc(t, df, hi)
               If temp > prob Then
                  ncp_t_nc1 = [#VALUE!]
                  Exit Function
               Else
                  checked_tnc_limit = True
               End If
               dif = dif * 1.99999999
            End If
         End If
         ncp = ncp + dif
      End If
   Loop While ((Abs(pr - prob) > prob * 0.00000000000001) And (Abs(dif) > Abs(ncp) * 0.0000000001))
   ncp_t_nc1 = ncp
   'Debug.Print "ncp_t_nc1", ncp_t_nc1
End Function

Private Function comp_ncp_t_nc1(ByVal prob As Double, ByVal t As Double, ByVal df As Double) As Double
'Uses Normal approx for non-central t (A&S 26.7.10) to get initial estimate the modified NR to improve it.
Dim ncp As Double, pr As Double, dif As Double, temp As Double, temp1 As Double, temp2 As Double, deriv As Double, checked_tnc_limit As Boolean, checked_0_limit As Boolean
Dim hi As Double, lo As Double, tnc_limit As Double, x As Double, y As Double
   If (prob > 0.5) Then
      comp_ncp_t_nc1 = ncp_t_nc1(1# - prob, t, df)
      Exit Function
   End If

   lo = 0#
   tnc_limit = Sqr(2# * nc_limit)
   hi = tnc_limit
   checked_0_limit = False
   checked_tnc_limit = False
   If t >= Min(1#, df) Then
      y = df / t
      x = t + y
      y = y / x
      x = t / x
   Else
      x = t * t
      y = df + x
      x = x / y
      y = df / y
   End If
   temp = -inv_normal(prob)
   temp1 = t * (1# - 0.25 / df)
   If t > df Then
        temp2 = temp * Sqr(t) * Sqr((1# / t + 0.5 * t / df))
   Else
        temp2 = temp * Sqr((1# + (0.5 * t / df) * t))
   End If
   ncp = Max(temp, temp1 + temp2)
   'Debug.Print "comp_ncp ncp estimate1", ncp
   If x > 1E-200 Then 'I think we can put more accurate bounds on when this will not deliver a sensible answer
      temp = invcompgamma(0.5 * x * df, prob) / y - 0.5
      If temp > 0 Then
         temp = Sqr(2# * temp)
         If temp > ncp Then
            temp = invgamma(0.5 * x * df, prob) / y - 0.5
            If temp > 0 Then
               ncp = Sqr(2# * temp)
            Else
               ncp = 0
            End If
         Else
            ncp = temp1 - temp2
         End If
      Else
         ncp = temp1 - temp2
      End If
   Else
      ncp = temp1 - temp2
   End If
   ncp = Min(Max(0#, ncp), tnc_limit)
   If ncp = 0# Then
      pr = comp_cdf_t_nc(t, df, 0#)
      If pr > prob Then
         If -invtdist(prob, df) <= t Then
            comp_ncp_t_nc1 = 0#
         Else
            comp_ncp_t_nc1 = [#VALUE!]
         End If
         Exit Function
      ElseIf Abs(pr - prob) <= -prob * 0.00000000000001 * Log(pr) Then
         comp_ncp_t_nc1 = 0#
         Exit Function
      Else
         checked_0_limit = True
      End If
      If x < y Then
         deriv = -comp_cdf_beta(x, 1, 0.5 * df) * OneOverSqrTwoPi
      Else
         deriv = -cdf_beta(y, 0.5 * df, 1) * OneOverSqrTwoPi
      End If
      If deriv = 0# Then
         ncp = tnc_limit
      Else
         ncp = (pr - prob) / deriv
         If ncp >= tnc_limit Then
            ncp = (pr / deriv) * logdif(pr, prob) 'If these two are miles apart then best to take invgamma estimate if > 0
         End If
      End If
   End If
   ncp = Min(ncp, tnc_limit)
   If ncp = tnc_limit Then
      pr = comp_cdf_t_nc(t, df, ncp)
      If pr < prob Then
         comp_ncp_t_nc1 = [#VALUE!]
         Exit Function
      Else
         checked_tnc_limit = True
      End If
   End If
   dif = ncp
   Do
      pr = comp_cdf_t_nc(t, df, ncp)
      'Debug.Print ncp, pr, prob
      If ncp > 1 Then
         deriv = comp_cdf_t_nc(t, df, ncp * (1 - 0.000001))
         deriv = 1000000# * (pr - deriv) / ncp
      ElseIf ncp > 0.000001 Then
         deriv = comp_cdf_t_nc(t, df, ncp + 0.000001)
         deriv = 1000000# * (deriv - pr)
      ElseIf x < y Then
         deriv = comp_cdf_beta(x, 1, 0.5 * df) * OneOverSqrTwoPi
      Else
         deriv = cdf_beta(y, 0.5 * df, 1) * OneOverSqrTwoPi
      End If
      If pr < 3E-308 And deriv = 0# Then
         lo = ncp
         dif = dif / 2#
         ncp = ncp - dif
      ElseIf deriv = 0# Then
         hi = ncp
         dif = dif / 2#
         ncp = ncp - dif
      Else
         If pr > prob Then
            hi = ncp
         Else
            lo = ncp
         End If
         dif = -(pr / deriv) * logdif(pr, prob)
         If ncp + dif < lo Then
            dif = (lo - ncp) / 2#
            If Not checked_0_limit And (lo = 0#) Then
               temp = comp_cdf_t_nc(t, df, lo)
               If temp > prob Then
                  If -invtdist(prob, df) <= t Then
                     comp_ncp_t_nc1 = 0#
                  Else
                     comp_ncp_t_nc1 = [#VALUE!]
                  End If
                  Exit Function
               Else
                  checked_0_limit = True
               End If
               dif = dif * 1.99999999
            End If
         ElseIf ncp + dif > hi Then
            dif = (hi - ncp) / 2#
            If Not checked_tnc_limit And (hi = tnc_limit) Then
               temp = comp_cdf_t_nc(t, df, hi)
               If temp < prob Then
                  comp_ncp_t_nc1 = [#VALUE!]
                  Exit Function
               Else
                  checked_tnc_limit = True
               End If
               dif = dif * 1.99999999
            End If
         End If
         ncp = ncp + dif
      End If
   Loop While ((Abs(pr - prob) > prob * 0.00000000000001) And (Abs(dif) > Abs(ncp) * 0.0000000001))
   comp_ncp_t_nc1 = ncp
   'Debug.Print "comp_ncp_t_nc1", comp_ncp_t_nc1
End Function

Public Function pdf_t_nc(ByVal x As Double, ByVal df As Double, ByVal nc_param As Double) As Double
'// Calculate pdf of noncentral t
'// Deliberately set not to calculate when x and nc_param have opposite signs as the algorithm used is prone to cancellation error in these circumstances.
'// The user can access t_nc1,comp_t_nc1 directly and check on the accuracy of the results, if required
  Dim nc_derivative As Double
  df = AlterForIntegralChecks_df(df)
  If (x < 0#) And (nc_param <= 0#) Then
     pdf_t_nc = pdf_t_nc(-x, df, -nc_param)
  ElseIf (df <= 0#) Or (nc_param < 0#) Or (nc_param > Sqr(2# * nc_limit)) Then
     pdf_t_nc = [#VALUE!]
  ElseIf (x < 0#) Then
     pdf_t_nc = [#VALUE!]
  ElseIf (x = 0# Or nc_param = 0#) Then
     pdf_t_nc = Exp(-nc_param ^ 2 / 2) * pdftdist(x, df)
  Else
     If (df < 1# Or x < 1# Or x <= nc_param) Then
        pdf_t_nc = t_nc1(x, df, nc_param, nc_derivative)
     Else
        pdf_t_nc = comp_t_nc1(x, df, nc_param, nc_derivative)
     End If
     If nc_derivative < cSmall Then
        pdf_t_nc = Exp(-nc_param ^ 2 / 2) * pdftdist(x, df)
     ElseIf df > 2# Then
        pdf_t_nc = nc_derivative / x
     Else
        pdf_t_nc = nc_derivative * (df / (2# * x))
     End If
  End If
  pdf_t_nc = GetRidOfMinusZeroes(pdf_t_nc)
End Function

Public Function cdf_t_nc(ByVal x As Double, ByVal df As Double, ByVal nc_param As Double) As Double
'// Calculate cdf of noncentral t
'// Deliberately set not to calculate when x and nc_param have opposite signs as the algorithm used is prone to cancellation error in these circumstances.
'// The user can access t_nc1,comp_t_nc1 directly and check on the accuracy of the results, if required
  Dim tdistDensity As Double, nc_derivative As Double
  df = AlterForIntegralChecks_df(df)
  If (nc_param = 0#) Then
     cdf_t_nc = tdist(x, df, tdistDensity)
  ElseIf (x <= 0#) And (nc_param < 0#) Then
     cdf_t_nc = comp_cdf_t_nc(-x, df, -nc_param)
  ElseIf (df <= 0#) Or (nc_param < 0#) Or (nc_param > Sqr(2# * nc_limit)) Then
     cdf_t_nc = [#VALUE!]
  ElseIf (x < 0#) Then
     cdf_t_nc = [#VALUE!]
  ElseIf (df < 1# Or x < 1# Or x <= nc_param) Then
     cdf_t_nc = t_nc1(x, df, nc_param, nc_derivative)
  Else
     cdf_t_nc = 1# - comp_t_nc1(x, df, nc_param, nc_derivative)
  End If
  cdf_t_nc = GetRidOfMinusZeroes(cdf_t_nc)
End Function

Public Function comp_cdf_t_nc(ByVal x As Double, ByVal df As Double, ByVal nc_param As Double) As Double
'// Calculate 1-cdf of noncentral t
'// Deliberately set not to calculate when x and nc_param have opposite signs as the algorithm used is prone to cancellation error in these circumstances.
'// The user can access t_nc1,comp_t_nc1 directly and check on the accuracy of the results, if required
  Dim tdistDensity As Double, nc_derivative As Double
  df = AlterForIntegralChecks_df(df)
  If (nc_param = 0#) Then
     comp_cdf_t_nc = tdist(-x, df, tdistDensity)
  ElseIf (x <= 0#) And (nc_param < 0#) Then
     comp_cdf_t_nc = cdf_t_nc(-x, df, -nc_param)
  ElseIf (df <= 0#) Or (nc_param < 0#) Or (nc_param > Sqr(2# * nc_limit)) Then
     comp_cdf_t_nc = [#VALUE!]
  ElseIf (x < 0#) Then
     comp_cdf_t_nc = [#VALUE!]
  ElseIf (df < 1# Or x < 1# Or x >= nc_param) Then
     comp_cdf_t_nc = comp_t_nc1(x, df, nc_param, nc_derivative)
  Else
     comp_cdf_t_nc = 1# - t_nc1(x, df, nc_param, nc_derivative)
  End If
  comp_cdf_t_nc = GetRidOfMinusZeroes(comp_cdf_t_nc)
End Function

Public Function inv_t_nc(ByVal prob As Double, ByVal df As Double, ByVal nc_param As Double) As Double
  df = AlterForIntegralChecks_df(df)
  If (nc_param = 0#) Then
     inv_t_nc = invtdist(prob, df)
  ElseIf (nc_param < 0#) Then
     inv_t_nc = -comp_inv_t_nc(prob, df, -nc_param)
  ElseIf (df <= 0# Or nc_param > Sqr(2# * nc_limit) Or prob <= 0# Or prob >= 1#) Then
     inv_t_nc = [#VALUE!]
  ElseIf (invcnormal(prob) < -nc_param) Then
     inv_t_nc = [#VALUE!]
  Else
     Dim oneMinusP As Double
     inv_t_nc = inv_t_nc1(prob, df, nc_param, oneMinusP)
  End If
  inv_t_nc = GetRidOfMinusZeroes(inv_t_nc)
End Function

Public Function comp_inv_t_nc(ByVal prob As Double, ByVal df As Double, ByVal nc_param As Double) As Double
  df = AlterForIntegralChecks_df(df)
  If (nc_param = 0#) Then
     comp_inv_t_nc = -invtdist(prob, df)
  ElseIf (nc_param < 0#) Then
     comp_inv_t_nc = -inv_t_nc(prob, df, -nc_param)
  ElseIf (df <= 0# Or nc_param > Sqr(2# * nc_limit) Or prob <= 0# Or prob >= 1#) Then
     comp_inv_t_nc = [#VALUE!]
  ElseIf (invcnormal(prob) > nc_param) Then
     comp_inv_t_nc = [#VALUE!]
  Else
     Dim oneMinusP As Double
     comp_inv_t_nc = comp_inv_t_nc1(prob, df, nc_param, oneMinusP)
  End If
  comp_inv_t_nc = GetRidOfMinusZeroes(comp_inv_t_nc)
End Function

Public Function ncp_t_nc(ByVal prob As Double, ByVal x As Double, ByVal df As Double) As Double
  df = AlterForIntegralChecks_df(df)
  If (x = 0# And prob > 0.5) Then
     ncp_t_nc = -invcnormal(prob)
  ElseIf (x < 0) Then
     ncp_t_nc = -comp_ncp_t_nc(prob, -x, df)
  ElseIf (df <= 0# Or prob <= 0# Or prob >= 1#) Then
     ncp_t_nc = [#VALUE!]
  Else
     ncp_t_nc = ncp_t_nc1(prob, x, df)
  End If
  ncp_t_nc = GetRidOfMinusZeroes(ncp_t_nc)
End Function

Public Function comp_ncp_t_nc(ByVal prob As Double, ByVal x As Double, ByVal df As Double) As Double
  df = AlterForIntegralChecks_df(df)
  If (x = 0#) Then
     comp_ncp_t_nc = invcnormal(prob)
  ElseIf (x < 0) Then
     comp_ncp_t_nc = -ncp_t_nc(prob, -x, df)
  ElseIf (df <= 0# Or prob <= 0# Or prob >= 1#) Then
     comp_ncp_t_nc = [#VALUE!]
  Else
     comp_ncp_t_nc = comp_ncp_t_nc1(prob, x, df)
  End If
  comp_ncp_t_nc = GetRidOfMinusZeroes(comp_ncp_t_nc)
End Function

Public Function pmf_GammaPoisson(i As Double, gamma_shape As Double, gamma_scale As Double) As Double
   Dim p As Double, q As Double, dfm As Double
   q = gamma_scale / (1# + gamma_scale)
   p = 1# / (1# + gamma_scale)
   i = AlterForIntegralChecks_Others(i)
   If (gamma_shape <= 0# Or gamma_scale <= 0#) Then
      pmf_GammaPoisson = [#VALUE!]
   ElseIf (i < 0#) Then
      pmf_GammaPoisson = 0
   Else
      If p < q Then
         dfm = gamma_shape - (gamma_shape + i) * p
      Else
         dfm = (gamma_shape + i) * q - i
      End If
      pmf_GammaPoisson = (gamma_shape / (gamma_shape + i)) * binomialTerm(i, gamma_shape, q, p, dfm, 0#)
   End If
   pmf_GammaPoisson = GetRidOfMinusZeroes(pmf_GammaPoisson)
End Function

Public Function cdf_GammaPoisson(i As Double, gamma_shape As Double, gamma_scale As Double) As Double
   Dim p As Double, q As Double
   q = gamma_scale / (1# + gamma_scale)
   p = 1# / (1# + gamma_scale)
   i = Int(i)
   If (gamma_shape <= 0# Or gamma_scale <= 0#) Then
      cdf_GammaPoisson = [#VALUE!]
   ElseIf (i < 0#) Then
      cdf_GammaPoisson = 0#
   ElseIf (p <= q) Then
      cdf_GammaPoisson = beta(p, gamma_shape, i + 1#)
   Else
      cdf_GammaPoisson = compbeta(q, i + 1#, gamma_shape)
   End If
   cdf_GammaPoisson = GetRidOfMinusZeroes(cdf_GammaPoisson)
End Function

Public Function comp_cdf_GammaPoisson(i As Double, gamma_shape As Double, gamma_scale As Double) As Double
   Dim p As Double, q As Double
   q = gamma_scale / (1# + gamma_scale)
   p = 1# / (1# + gamma_scale)
   i = Int(i)
   If (gamma_shape <= 0# Or gamma_scale <= 0#) Then
      comp_cdf_GammaPoisson = [#VALUE!]
   ElseIf (i < 0#) Then
      comp_cdf_GammaPoisson = 1#
   ElseIf (p <= q) Then
      comp_cdf_GammaPoisson = compbeta(p, gamma_shape, i + 1#)
   Else
      comp_cdf_GammaPoisson = beta(q, i + 1#, gamma_shape)
   End If
   comp_cdf_GammaPoisson = GetRidOfMinusZeroes(comp_cdf_GammaPoisson)
End Function

Public Function crit_GammaPoisson(gamma_shape As Double, gamma_scale As Double, crit_prob As Double) As Double
   Dim p As Double, q As Double
   q = gamma_scale / (1# + gamma_scale)
   p = 1# / (1# + gamma_scale)
   If (gamma_shape < 0# Or gamma_scale < 0#) Then
      crit_GammaPoisson = [#VALUE!]
   ElseIf (crit_prob < 0# Or crit_prob >= 1#) Then
      crit_GammaPoisson = [#VALUE!]
   ElseIf (crit_prob = 0#) Then
      crit_GammaPoisson = [#VALUE!]
   Else
      Dim i As Double, pr As Double
      crit_GammaPoisson = critnegbinom(gamma_shape, p, q, crit_prob)
      i = crit_GammaPoisson
      If p <= q Then
         pr = beta(p, gamma_shape, i + 1#)
      Else
         pr = compbeta(q, i + 1#, gamma_shape)
      End If
      If (pr = crit_prob) Then
      ElseIf (pr > crit_prob) Then
         i = i - 1#
         If p <= q Then
            pr = beta(p, gamma_shape, i + 1#)
         Else
            pr = compbeta(q, i + 1#, gamma_shape)
         End If
         If (pr >= crit_prob) Then
            crit_GammaPoisson = i
         End If
      Else
         crit_GammaPoisson = i + 1#
      End If
   End If
   crit_GammaPoisson = GetRidOfMinusZeroes(crit_GammaPoisson)
End Function

Public Function comp_crit_GammaPoisson(gamma_shape As Double, gamma_scale As Double, crit_prob As Double) As Double
   Dim p As Double, q As Double
   q = gamma_scale / (1# + gamma_scale)
   p = 1# / (1# + gamma_scale)
   If (gamma_shape < 0# Or gamma_scale < 0#) Then
      comp_crit_GammaPoisson = [#VALUE!]
   ElseIf (crit_prob <= 0# Or crit_prob > 1#) Then
      comp_crit_GammaPoisson = [#VALUE!]
   ElseIf (crit_prob = 1#) Then
      comp_crit_GammaPoisson = [#VALUE!]
   Else
      Dim i As Double, pr As Double
      comp_crit_GammaPoisson = critcompnegbinom(gamma_shape, p, q, crit_prob)
      i = comp_crit_GammaPoisson
      If p <= q Then
         pr = compbeta(p, gamma_shape, i + 1#)
      Else
         pr = beta(q, i + 1#, gamma_shape)
      End If
      If (pr = crit_prob) Then
      ElseIf (pr < crit_prob) Then
         i = i - 1#
         If p <= q Then
            pr = compbeta(p, gamma_shape, i + 1#)
         Else
            pr = beta(q, i + 1#, gamma_shape)
         End If
         If (pr <= crit_prob) Then
            comp_crit_GammaPoisson = i
         End If
      Else
         comp_crit_GammaPoisson = i + 1#
      End If
   End If
   comp_crit_GammaPoisson = GetRidOfMinusZeroes(comp_crit_GammaPoisson)
End Function

Private Function PBB(ByVal i As Double, ByVal ssmi As Double, ByVal beta_shape1 As Double, ByVal beta_shape2 As Double) As Double
    hTerm = hypergeometricTerm(i, ssmi, beta_shape1, beta_shape2)
    PBB = (beta_shape1 / (i + beta_shape1)) * (beta_shape2 / (beta_shape1 + beta_shape2)) * ((i + ssmi + beta_shape1 + beta_shape2) / (ssmi + beta_shape2)) * hTerm
End Function

Private Function PBNB(ByVal i As Double, ByVal r As Double, ByVal beta_shape1 As Double, ByVal beta_shape2 As Double) As Double
    hTerm = hypergeometricTerm(i, r, beta_shape2, beta_shape1)
    PBNB = (beta_shape2 / (beta_shape1 + beta_shape2)) * (r / (beta_shape1 + r)) * beta_shape1 * (i + beta_shape1 + r + beta_shape2) / ((i + r) * (i + beta_shape2)) * hTerm
End Function

Public Function pmf_BetaNegativeBinomial(ByVal i As Double, ByVal r As Double, ByVal beta_shape1 As Double, ByVal beta_shape2 As Double) As Double
   i = AlterForIntegralChecks_Others(i)
   If (r <= 0# Or beta_shape1 <= 0# Or beta_shape2 <= 0#) Then
      pmf_BetaNegativeBinomial = [#VALUE!]
   ElseIf i < 0 Then
      pmf_BetaNegativeBinomial = 0#
   Else
      pmf_BetaNegativeBinomial = (beta_shape2 / (beta_shape1 + beta_shape2)) * (r / (beta_shape1 + r)) * beta_shape1 * (i + beta_shape1 + r + beta_shape2) / ((i + r) * (i + beta_shape2)) * hypergeometricTerm(i, r, beta_shape2, beta_shape1)
   End If
   pmf_BetaNegativeBinomial = GetRidOfMinusZeroes(pmf_BetaNegativeBinomial)
End Function

Private Function CBNB0(ByVal i As Double, ByVal r As Double, ByVal beta_shape1 As Double, ByVal beta_shape2 As Double, toBeAdded As Double) As Double
   Dim ha1 As Double, hprob As Double, hswap As Boolean
   Dim mrb2 As Double, other As Double, temp As Double
   If (r < 2# Or beta_shape2 < 2#) Then
'One assumption here that i is integral or greater than 4.
      mrb2 = Max(r, beta_shape2)
      other = Min(r, beta_shape2)
      CBNB0 = PBB(i, other, mrb2, beta_shape1)
      If i = 0# Then Exit Function
      CBNB0 = CBNB0 * (1# + i * (other + beta_shape1) / (((i - 1#) + mrb2) * (other + 1#)))
      If i = 1# Then Exit Function
      i = i - 2#
      other = other + 2#
      temp = PBB(i, mrb2, other, beta_shape1)
      If i = 0# Then
         CBNB0 = CBNB0 + temp
         Exit Function
      End If
      CBNB0 = CBNB0 + temp * (1# + i * (mrb2 + beta_shape1) / (((i - 1#) + other) * (mrb2 + 1#)))
      If i = 1# Then Exit Function
      i = i - 2#
      mrb2 = mrb2 + 2#
      CBNB0 = CBNB0 + CBNB0(i, mrb2, beta_shape1, other, CBNB0)
   ElseIf (beta_shape1 < 1#) Then
      mrb2 = Max(r, beta_shape2)
      other = Min(r, beta_shape2)
      CBNB0 = hypergeometric(i, mrb2 - 1#, other, beta_shape1, False, ha1, hprob, hswap)
      If hswap Then
         temp = PBB(mrb2 - 1#, beta_shape1, i + 1#, other)
         If (toBeAdded + (CBNB0 - temp)) < 0.01 * (toBeAdded + (CBNB0 + temp)) Then
            CBNB0 = CBNB2(i, mrb2, beta_shape1, other)
         Else
            CBNB0 = CBNB0 - temp
         End If
      ElseIf ha1 < -0.9 * beta_shape1 / (beta_shape1 + other) Then
         CBNB0 = [#VALUE!]
      Else
         CBNB0 = hprob * (beta_shape1 / (beta_shape1 + other) + ha1)
      End If
   Else
      CBNB0 = hypergeometric(i, r, beta_shape2, beta_shape1 - 1#, False, ha1, hprob, hswap)
   End If
End Function

Private Function CBNB2(ByVal i As Double, ByVal r As Double, ByVal beta_shape1 As Double, ByVal beta_shape2 As Double) As Double
   Dim j As Double, ss As Double, bs2 As Double, temp As Double, d1 As Double, d2 As Double, d_count As Double, pbbval As Double
   'In general may be a good idea to take Min(i, beta_shape1) down to just above 0 and then work on Max(i, beta_shape1)
   ss = Min(r, beta_shape2)
   bs2 = Max(r, beta_shape2)
   r = ss
   beta_shape2 = bs2
   d1 = (i + 0.5) * (beta_shape1 + 0.5) - (bs2 - 1.5) * (ss - 0.5)
   If d1 < 0# Then
      CBNB2 = CBNB0(i, ss, beta_shape1, bs2, 0#)
      Exit Function
   End If
   d1 = Int(d1 / (bs2 + beta_shape1 - 1#)) + 10#
   If ss + d1 > bs2 Then d1 = Int(bs2 - ss)
   ss = ss + d1
   j = i - d1
   d2 = (j + 0.5) * (beta_shape1 + 0.5) - (bs2 - 1.5) * (ss - 0.5)
   If d2 < 0# Then
      d2 = 10#
   Else
      temp = bs2 + ss + 2# * beta_shape1 - 1#
      d2 = Int((Sqr(temp ^ 2 + 4# * d2) - temp) / 2#) + 10#
   End If
   If 2# * d2 > i Then
      d2 = Int(i / 2#)
   End If
   pbbval = PBB(i, r, beta_shape2, beta_shape1)
   ss = ss + d2
   bs2 = bs2 + d2
   j = j - 2# * d2
   CBNB2 = CBNB0(j, ss, beta_shape1, bs2, 0#)
   temp = 1#
   d_count = d2 - 2#
   j = j + 1#
   Do While d_count >= 0#
      j = j + 1#
      bs2 = beta_shape2 + d_count
      d_count = d_count - 1#
      temp = 1# + (j * (bs2 + beta_shape1) / ((j + ss - 1#) * (bs2 + 1#))) * temp
   Loop
   j = i - d2 - d1
   temp = (ss * (j + bs2)) / (bs2 * (j + ss)) * temp
   d_count = d1 + d2 - 1#
   Do While d_count >= 0
      j = j + 1#
      ss = r + d_count
      d_count = d_count - 1#
      temp = 1# + (j * (ss + beta_shape1) / ((j + bs2 - 1#) * (ss + 1#))) * temp
   Loop
   CBNB2 = CBNB2 + temp * pbbval
   Exit Function
End Function

Public Function cdf_BetaNegativeBinomial(ByVal i As Double, ByVal r As Double, ByVal beta_shape1 As Double, ByVal beta_shape2 As Double) As Double
   i = Int(i)
   If (r <= 0# Or beta_shape1 <= 0# Or beta_shape2 <= 0#) Then
      cdf_BetaNegativeBinomial = [#VALUE!]
   ElseIf i < 0 Then
      cdf_BetaNegativeBinomial = 0#
   Else
      cdf_BetaNegativeBinomial = CBNB0(i, r, beta_shape1, beta_shape2, 0#)
   End If
   cdf_BetaNegativeBinomial = GetRidOfMinusZeroes(cdf_BetaNegativeBinomial)
End Function

Public Function comp_cdf_BetaNegativeBinomial(ByVal i As Double, ByVal r As Double, ByVal beta_shape1 As Double, ByVal beta_shape2 As Double) As Double
   Dim ha1 As Double, hprob As Double, hswap As Boolean
   Dim mrb2 As Double, other As Double, temp As Double, mnib1 As Double, mxib1 As Double, swap As Double, max_iterations As Double
   i = Int(i)
   mrb2 = Max(r, beta_shape2)
   other = Min(r, beta_shape2)
   If (other <= 0# Or beta_shape1 <= 0#) Then
      comp_cdf_BetaNegativeBinomial = [#VALUE!]
   ElseIf i < 0# Then
      comp_cdf_BetaNegativeBinomial = 1#
   ElseIf (i = 0#) Or ((i < 1000000#) And (other < 0.001) And (beta_shape1 > 50# * other) And (100# * i * beta_shape1 < mrb2)) Then
      comp_cdf_BetaNegativeBinomial = ccBNB5(i, mrb2, beta_shape1, other)
   ElseIf (mrb2 >= 100# Or other > 20# Or (mrb2 >= 5# And (other - 0.5) * (mrb2 - 0.5) > (i + 0.5) * (beta_shape1 + 0.5))) Then
      comp_cdf_BetaNegativeBinomial = CBNB0(mrb2 - 1#, i + 1#, other, beta_shape1, 0#)
   Else
      comp_cdf_BetaNegativeBinomial = 0#
      temp = 0#
      i = i + 1#
      If other >= 1# Then
         mrb2 = mrb2 - 1#
         other = other - 1#
         temp = hypergeometricTerm(i, mrb2, other, beta_shape1)
         comp_cdf_BetaNegativeBinomial = temp
         Do While (other >= 1#) And (temp > 1E-16 * comp_cdf_BetaNegativeBinomial)
            i = i + 1#
            beta_shape1 = beta_shape1 + 1#
            temp = temp * (mrb2 * other) / (i * beta_shape1)
            mrb2 = mrb2 - 1#
            other = other - 1#
            comp_cdf_BetaNegativeBinomial = comp_cdf_BetaNegativeBinomial + temp
         Loop
         If other >= 1# Then Exit Function
         i = i + 1#
         beta_shape1 = beta_shape1 + 1#
      End If
      If mrb2 >= 1# Then
         mxib1 = Max(i, beta_shape1)
         mnib1 = Min(i, beta_shape1)
         If temp = 0# Then
            mrb2 = mrb2 - 1#
            temp = PBB(mnib1, mrb2, other, mxib1)
         Else 'temp is hypergeometricTerm(mnib1-1, mrb2, other, mxib1-1)
            temp = temp * other * mrb2
            mrb2 = mrb2 - 1#
            temp = temp / (mnib1 * (mrb2 + mxib1))
         End If
         comp_cdf_BetaNegativeBinomial = comp_cdf_BetaNegativeBinomial + temp
         Do While (mrb2 >= 1#) And (temp > 1E-16 * comp_cdf_BetaNegativeBinomial)
            temp = temp * mrb2 * (mnib1 + other)
            mnib1 = mnib1 + 1#
            If mnib1 > mxib1 Then
               swap = mxib1
               mxib1 = mnib1
               mnib1 = swap
            End If
'Block below not required if hypergeometric block included above and therefore other guaranteed < 1 <= mrb2
            'If mrb2 < other Then
            '   swap = other
            '   other = mrb2
            '   mrb2 = swap
            'End If
            mrb2 = mrb2 - 1#
            temp = temp / ((mrb2 + mxib1) * mnib1)
            comp_cdf_BetaNegativeBinomial = comp_cdf_BetaNegativeBinomial + temp
         Loop
         If mrb2 >= 1# Then Exit Function
         temp = temp * mrb2 / (mnib1 + mrb2)
      Else
         mxib1 = beta_shape1
         mnib1 = i
         If temp = 0# Then
            temp = pBNB(mnib1, mrb2, mxib1, other)
         Else
            temp = temp * mrb2 * other / (i * (mrb2 + other + mnib1 + mxib1 + -1))
         End If
         comp_cdf_BetaNegativeBinomial = comp_cdf_BetaNegativeBinomial + temp
      End If
      max_iterations = 60#
      Do
         temp = temp * (mnib1 + mrb2) * (mnib1 + other) / (mnib1 + mxib1 + mrb2 + other)
         mnib1 = mnib1 + 1#
         If mxib1 < mnib1 Then
            swap = mxib1
            mxib1 = mnib1
            mnib1 = swap
         End If
         temp = temp / mnib1
         comp_cdf_BetaNegativeBinomial = comp_cdf_BetaNegativeBinomial + temp
      Loop Until (temp <= 1E-16 * comp_cdf_BetaNegativeBinomial) Or (mnib1 + mxib1 > max_iterations)
      temp = temp * (mnib1 + mrb2) * (mnib1 + other) / ((mnib1 + 1#) * mxib1)
      comp_cdf_BetaNegativeBinomial = comp_cdf_BetaNegativeBinomial + temp
      mnib1 = mnib1 + 1#
      mrb2 = mrb2 - 1#
      other = other - 1#
      Do
         mnib1 = mnib1 + 1#
         mxib1 = mxib1 + 1#
         temp = temp * (mrb2 * other) / (mnib1 * mxib1)
         mrb2 = mrb2 - 1#
         other = other - 1#
         comp_cdf_BetaNegativeBinomial = comp_cdf_BetaNegativeBinomial + temp
      Loop Until Abs(temp) <= 1E-16 * comp_cdf_BetaNegativeBinomial
   End If
   comp_cdf_BetaNegativeBinomial = GetRidOfMinusZeroes(comp_cdf_BetaNegativeBinomial)
End Function

Private Function critbetanegbinomial(ByVal a As Double, ByVal b As Double, ByVal r As Double, ByVal cprob As Double) As Double
'//i such that Pr(betanegbinomial(i,r,a,b)) >= cprob and  Pr(betanegbinomial(i-1,r,a,b)) < cprob
   If (cprob > 0.5) Then
      critbetanegbinomial = critcompbetanegbinomial(a, b, r, 1# - cprob)
      Exit Function
   End If
   Dim pr As Double, tpr As Double
   Dim i As Double, temp As Double, temp2 As Double, oneMinusP As Double
   If b > r Then
      i = b
      b = r
      r = i
   End If
   If (a < 10# Or b < 10#) Then
      If r < a And a < 1# Then
         pr = cprob * a / r
      Else
         pr = cprob
      End If
      i = invcompbeta(a, b, pr, oneMinusP)
   Else
      pr = r / (r + a + b - 1#)
      i = invcompbeta(a * pr, b * pr, cprob, oneMinusP)
   End If
   If i = 0# Then
      i = max_crit / 2#
   Else
      i = r * (oneMinusP / i)
      If i >= max_crit Then
         i = max_crit - 1#
      End If
   End If
   While (True)
      If (i < 0#) Then
         i = 0#
      End If
      i = Int(i + 0.5)
      If (i >= max_crit) Then
         critbetanegbinomial = [#VALUE!]
         Exit Function
      End If
      pr = CBNB0(i, r, a, b, 0#)
      tpr = 0#
      If (pr > cprob * (1 + cfSmall)) Then
         If (i = 0#) Then
            critbetanegbinomial = 0#
            Exit Function
         End If
         tpr = pmf_BetaNegativeBinomial(i, r, a, b)
         If (pr < (1# + 0.00001) * tpr) Then
            i = i - 1#
            tpr = tpr * (((i + 1#) * (i + a + r + b)) / ((i + r) * (i + b)))
            While (tpr > cprob)
               i = i - 1#
               tpr = tpr * (((i + 1#) * (i + a + r + b)) / ((i + r) * (i + b)))
            Wend
         Else
            pr = pr - tpr
            If (pr < cprob) Then
               critbetanegbinomial = i
               Exit Function
            End If
            i = i - 1#
            If (i = 0#) Then
               critbetanegbinomial = 0#
               Exit Function
            End If
            temp = (pr - cprob) / tpr
            If (temp > 10#) Then
               temp = Int(temp + 0.5)
               If temp > i Then
                  i = i / 10#
               Else
                  i = Int(i - temp)
                  temp2 = pmf_BetaNegativeBinomial(i, r, a, b)
                  i = i - temp * (tpr - temp2) / (2# * temp2)
               End If
            Else
               tpr = tpr * (((i + 1#) * (i + a + r + b)) / ((i + r) * (i + b)))
               pr = pr - tpr
               If (pr < cprob) Then
                  critbetanegbinomial = i
                  Exit Function
               End If
               temp2 = (pr - cprob) / tpr
               If (temp2 < temp - 0.9) Then
                  While (pr >= cprob)
                     i = i - 1#
                     tpr = tpr * (((i + 1#) * (i + a + r + b)) / ((i + r) * (i + b)))
                     pr = pr - tpr
                  Wend
                  critbetanegbinomial = i
                  Exit Function
               Else
                  temp = Int(Log(cprob / pr) / Log((((i + 1#) * (i + a + r + b)) / ((i + r) * (i + b)))) + 0.5)
                  i = i - temp
                  temp2 = pmf_BetaNegativeBinomial(i, r, a, b)
                  If (temp2 > nearly_zero) Then
                     temp = Log((cprob / pr) * (tpr / temp2)) / Log((((i + 1#) * (i + a + r + b)) / ((i + r) * (i + b))))
                     i = i - temp
                  End If
               End If
            End If
         End If
      ElseIf ((1# + cfSmall) * pr < cprob) Then
         While ((tpr < cSmall) And (pr < cprob))
            i = i + 1#
            tpr = pmf_BetaNegativeBinomial(i, r, a, b)
            pr = pr + tpr
            If pr = 0# Or 1E+100 * pr < cprob Then
               tpr = cSmall
            End If
         Wend
         If pr > 0# Then
            temp = (cprob - pr) / tpr
         Else
            temp = max_crit
         End If
         If temp <= 0# Then
            critbetanegbinomial = i
            Exit Function
         ElseIf temp < 10# Then
            While (pr < cprob)
               tpr = tpr * (((i + r) * (i + b)) / ((i + 1#) * (i + a + r + b)))
               pr = pr + tpr
               i = i + 1#
            Wend
            critbetanegbinomial = i
            Exit Function
         ElseIf i = max_crit Then
            critbetanegbinomial = [#VALUE!]
            Exit Function
         ElseIf i + temp > max_crit Then
            i = max_crit - 1
         Else
            i = Int(i + temp)
            temp2 = pmf_BetaNegativeBinomial(i, r, a, b)
            If temp2 > 0# Then i = i + temp * (tpr - temp2) / (2# * temp2)
         End If
      Else
         critbetanegbinomial = i
         Exit Function
      End If
   Wend
End Function

Private Function critcompbetanegbinomial(ByVal a As Double, ByVal b As Double, ByVal r As Double, ByVal cprob As Double) As Double
'//i such that 1-Pr(betanegbinomial(i,r,a,b)) > cprob and  1-Pr(betanegbinomial(i-1,r,a,b)) <= cprob
   If (cprob > 0.5) Then
      critcompbetanegbinomial = critbetanegbinomial(a, b, r, 1# - cprob)
      Exit Function
   End If
   Dim pr As Double, tpr As Double, i_smallest As Double
   Dim i As Double, temp As Double, temp2 As Double, oneMinusP As Double
   i_smallest = 0#
   If b > r Then
      i = b
      b = r
      r = i
   End If
   If (a < 10# Or b < 10#) Then
      If r < a And a < 1# Then
         pr = cprob * a / r
      Else
         pr = cprob
      End If
      i = invbeta(a, b, pr, oneMinusP)
   Else
      pr = r / (r + a + b - 1#)
      i = invbeta(a * pr, b * pr, cprob, oneMinusP)
   End If
   If i = 0# Then
      i = max_crit / 2#
   Else
      i = r * (oneMinusP / i)
      If i >= max_crit Then
         i = max_crit - 1#
      End If
   End If
   While (True)
      If (i < 0#) Then
         i = 0#
      End If
      i = Int(i + 0.5)
      If (i >= max_crit) Then
         critcompbetanegbinomial = [#VALUE!]
         Exit Function
      End If
      pr = comp_cdf_BetaNegativeBinomial(i, r, a, b)
      tpr = 0#
      If (pr > cprob * (1 + cfSmall)) Then
         i = i + 1#
         i_smallest = i
         tpr = pmf_BetaNegativeBinomial(i, r, a, b)
         If (pr < (1.00001) * tpr) Then
            While (tpr > cprob)
               tpr = tpr * (((i + r) * (i + b)) / ((i + 1#) * (i + a + r + b)))
               i = i + 1#
            Wend
         Else
            pr = pr - tpr
            If (pr <= cprob) Then
               critcompbetanegbinomial = i
               Exit Function
            End If
            temp = (pr - cprob) / tpr
            If (temp > 10#) Then
               temp = Int(temp + 0.5)
               i = i + temp
               temp2 = pmf_BetaNegativeBinomial(i, r, a, b)
               i = i + temp * (tpr - temp2) / (2# * temp2)
            Else
               tpr = tpr * (((i + r) * (i + b)) / ((i + 1#) * (i + a + r + b)))
               i = i + 1#
               pr = pr - tpr
               If (pr <= cprob) Then
                  critcompbetanegbinomial = i
                  Exit Function
               End If
               temp2 = (pr - cprob) / tpr
               If (temp2 < temp - 0.9) Then
                  While (pr > cprob)
                     tpr = tpr * (((i + r) * (i + b)) / ((i + 1#) * (i + a + r + b)))
                     i = i + 1#
                     pr = pr - tpr
                  Wend
                  critcompbetanegbinomial = i
                  Exit Function
               Else
                  temp = Int(Log(cprob / pr) / Log((((i + r - 1#) * (i + b - 1#)) / (i * (i + a + r + b - 1#)))) + 0.5)
                  i = i + temp
                  temp2 = pmf_BetaNegativeBinomial(i, r, a, b)
                  If (temp2 > nearly_zero) Then
                     temp = Log((cprob / pr) * (tpr / temp2)) / Log((((i + r - 1#) * (i + b - 1#)) / (i * (i + a + r + b - 1#))))
                     i = i + temp
                  End If
               End If
            End If
         End If
      ElseIf pr < excel0 Then
         i = (i_smallest + i) / 2
      ElseIf ((1# + cfSmall) * pr < cprob) Then
         While ((tpr < cSmall) And (pr <= cprob))
            tpr = pmf_BetaNegativeBinomial(i, r, a, b)
            pr = pr + tpr
            i = i - 1#
         Wend
         temp = (cprob - pr) / tpr
         If temp <= 0# Then
            critcompbetanegbinomial = i + 1#
            Exit Function
         ElseIf temp < 100# Or i < 1000# Then
            While (pr <= cprob)
               tpr = tpr * (((i + 1#) * (i + a + r + b)) / ((i + r) * (i + b)))
               pr = pr + tpr
               i = i - 1#
            Wend
            critcompbetanegbinomial = i + 1#
            Exit Function
         ElseIf temp > i Then
            i = i / 10#
         Else
            i = Int(i - temp)
            temp2 = pmf_BetaNegativeBinomial(i, r, a, b)
            If temp2 > 0# Then i = i - temp * (tpr - temp2) / (2# * temp2)
         End If
      Else
         critcompbetanegbinomial = i
         Exit Function
      End If
   Wend
End Function

Public Function crit_BetaNegativeBinomial(ByVal r As Double, ByVal beta_shape1 As Double, ByVal beta_shape2 As Double, ByVal crit_prob As Double) As Double
   If (beta_shape1 <= 0# Or beta_shape2 <= 0# Or r <= 0#) Then
      crit_BetaNegativeBinomial = [#VALUE!]
   ElseIf (crit_prob < 0# Or crit_prob >= 1#) Then
      crit_BetaNegativeBinomial = [#VALUE!]
   ElseIf (crit_prob = 0#) Then
      crit_BetaNegativeBinomial = [#VALUE!]
   Else
      Dim i As Double, pr As Double
      i = critbetanegbinomial(beta_shape1, beta_shape2, r, crit_prob)
      crit_BetaNegativeBinomial = i
      pr = cdf_BetaNegativeBinomial(i, r, beta_shape1, beta_shape2)
      If (pr = crit_prob) Then
      ElseIf (pr > crit_prob) Then
         i = i - 1#
         pr = cdf_BetaNegativeBinomial(i, r, beta_shape1, beta_shape2)
         If (pr >= crit_prob) Then
            crit_BetaNegativeBinomial = i
         End If
      Else
         crit_BetaNegativeBinomial = i + 1#
      End If
   End If
   crit_BetaNegativeBinomial = GetRidOfMinusZeroes(crit_BetaNegativeBinomial)
End Function

Public Function comp_crit_BetaNegativeBinomial(ByVal r As Double, ByVal beta_shape1 As Double, ByVal beta_shape2 As Double, ByVal crit_prob As Double) As Double
   If (beta_shape1 <= 0# Or beta_shape2 <= 0# Or r <= 0#) Then
      comp_crit_BetaNegativeBinomial = [#VALUE!]
   ElseIf (crit_prob <= 0# Or crit_prob > 1#) Then
      comp_crit_BetaNegativeBinomial = [#VALUE!]
   ElseIf (crit_prob = 1#) Then
      comp_crit_BetaNegativeBinomial = 0#
   Else
      Dim i As Double, pr As Double
      i = critcompbetanegbinomial(beta_shape1, beta_shape2, r, crit_prob)
      comp_crit_BetaNegativeBinomial = i
      pr = comp_cdf_BetaNegativeBinomial(i, r, beta_shape1, beta_shape2)
      If (pr = crit_prob) Then
      ElseIf (pr < crit_prob) Then
         i = i - 1#
         pr = comp_cdf_BetaNegativeBinomial(i, r, beta_shape1, beta_shape2)
         If (pr <= crit_prob) Then
            comp_crit_BetaNegativeBinomial = i
         End If
      Else
         comp_crit_BetaNegativeBinomial = i + 1#
      End If
   End If
   comp_crit_BetaNegativeBinomial = GetRidOfMinusZeroes(comp_crit_BetaNegativeBinomial)
End Function

Public Function pmf_BetaBinomial(ByVal i As Double, ByVal sample_size As Double, ByVal beta_shape1 As Double, ByVal beta_shape2 As Double) As Double
   i = AlterForIntegralChecks_Others(i)
   sample_size = AlterForIntegralChecks_Others(sample_size)
   If (beta_shape1 <= 0# Or beta_shape2 <= 0# Or sample_size < 0#) Then
      pmf_BetaBinomial = [#VALUE!]
   ElseIf i < 0 Or i > sample_size Then
      pmf_BetaBinomial = 0#
   Else
      pmf_BetaBinomial = (beta_shape1 / (i + beta_shape1)) * (beta_shape2 / (beta_shape1 + beta_shape2)) * ((sample_size + beta_shape1 + beta_shape2) / (sample_size - i + beta_shape2)) * hypergeometricTerm(i, sample_size - i, beta_shape1, beta_shape2)
   End If
   pmf_BetaBinomial = GetRidOfMinusZeroes(pmf_BetaBinomial)
End Function

Public Function cdf_BetaBinomial(ByVal i As Double, ByVal sample_size As Double, ByVal beta_shape1 As Double, ByVal beta_shape2 As Double) As Double
   i = Int(i)
   sample_size = AlterForIntegralChecks_Others(sample_size)
   If (beta_shape1 <= 0# Or beta_shape2 <= 0# Or sample_size < 0#) Then
      cdf_BetaBinomial = [#VALUE!]
   ElseIf i < 0# Then
      cdf_BetaBinomial = 0#
   Else
      i = i + 1#
      cdf_BetaBinomial = comp_cdf_BetaNegativeBinomial(sample_size - i, i, beta_shape1, beta_shape2)
   End If
End Function

Public Function comp_cdf_BetaBinomial(ByVal i As Double, ByVal sample_size As Double, ByVal beta_shape1 As Double, ByVal beta_shape2 As Double) As Double
   i = Int(i)
   sample_size = AlterForIntegralChecks_Others(sample_size)
   If (beta_shape1 <= 0# Or beta_shape2 <= 0# Or sample_size < 0#) Then
      comp_cdf_BetaBinomial = [#VALUE!]
   ElseIf i < 0# Then
      comp_cdf_BetaBinomial = 1#
   ElseIf i >= sample_size Then
      comp_cdf_BetaBinomial = 0#
   Else
      comp_cdf_BetaBinomial = comp_cdf_BetaNegativeBinomial(i, sample_size - i, beta_shape2, beta_shape1)
   End If
End Function

Private Function critbetabinomial(ByVal a As Double, ByVal b As Double, ByVal ss As Double, ByVal cprob As Double) As Double
'//i such that Pr(betabinomial(i,ss,a,b)) >= cprob and  Pr(betabinomial(i-1,ss,a,b)) < cprob
   If (cprob > 0.5) Then
      critbetabinomial = critcompbetabinomial(a, b, ss, 1# - cprob)
      Exit Function
   End If
   Dim pr As Double, tpr As Double
   Dim i As Double, temp As Double, temp2 As Double, oneMinusP As Double
   If (a + b < 1#) Then
      i = invbeta(a, b, cprob, oneMinusP) * ss
   Else
      pr = ss / (ss + a + b - 1#)
      i = invbeta(a * pr, b * pr, cprob, oneMinusP) * ss
   End If
   While (True)
      If (i < 0#) Then
         i = 0#
      ElseIf (i > ss) Then
         i = ss
      End If
      i = Int(i + 0.5)
      If (i >= max_discrete) Then
         critbetabinomial = i
         Exit Function
      End If
      pr = cdf_BetaBinomial(i, ss, a, b)
      tpr = 0#
      If (pr >= cprob * (1 + cfSmall)) Then
         If (i = 0#) Then
            critbetabinomial = 0#
            Exit Function
         End If
         tpr = pmf_BetaBinomial(i, ss, a, b)
         pr = pr - tpr
         If (pr < cprob) Then
            critbetabinomial = i
            Exit Function
         End If
         tpr = tpr * (i * ((ss - i) + b)) / ((a + i - 1#) * (ss - i + 1#))
         i = i - 1#
         If (pr < (1# + 0.00001) * tpr) Then
            While (tpr > cprob)
               tpr = tpr * (i * ((ss - i) + b)) / ((a + i - 1#) * (ss - i + 1#))
               i = i - 1#
            Wend
         Else
            If (i = 0#) Then
               critbetabinomial = 0#
               Exit Function
            End If
            temp = (pr - cprob) / tpr
            If (temp > 10) Then
               temp = Int(temp + 0.5)
               i = i - temp
               temp2 = pmf_BetaBinomial(i, ss, a, b)
               i = i - temp * (tpr - temp2) / (2# * temp2)
            Else
               tpr = tpr * (i * ((ss - i) + b)) / ((a + i - 1#) * (ss - i + 1#))
               pr = pr - tpr
               If (pr < cprob) Then
                  critbetabinomial = i
                  Exit Function
               End If
               i = i - 1#
               temp2 = (pr - cprob) / tpr
               If (temp2 < temp - 0.9) Then
                  While (pr >= cprob)
                     tpr = tpr * (i * ((ss - i) + b)) / ((a + i - 1#) * (ss - i + 1#))
                     pr = pr - tpr
                     i = i - 1#
                  Wend
                  critbetabinomial = i + 1#
                  Exit Function
               Else
                  temp = Int(Log(cprob / pr) / Log((i * ((ss - i) + b)) / ((a + i - 1#) * (ss - i + 1#))) + 0.5)
                  i = i - temp
                  temp2 = pmf_BetaBinomial(i, ss, a, b)
                  If (temp2 > nearly_zero) Then
                     temp = Log((cprob / pr) * (tpr / temp2)) / Log((i * ((ss - i) + b)) / ((a + i - 1#) * (ss - i + 1#)))
                     i = i - temp
                  End If
               End If
            End If
         End If
      ElseIf ((1# + cfSmall) * pr < cprob) Then
         While ((tpr < cSmall) And (pr < cprob))
            i = i + 1#
            tpr = pmf_BetaBinomial(i, ss, a, b)
            pr = pr + tpr
         Wend
         temp = (cprob - pr) / tpr
         If temp <= 0# Then
            critbetabinomial = i
            Exit Function
         ElseIf temp < 10# Then
            While (pr < cprob)
               i = i + 1#
               tpr = tpr * ((a + i - 1#) * (ss - i + 1#)) / (i * ((ss - i) + b))
               pr = pr + tpr
            Wend
            critbetabinomial = i
            Exit Function
         ElseIf temp > 4E+15 Then
            i = 4E+15
         Else
            i = Int(i + temp)
            temp2 = pmf_BetaBinomial(i, ss, a, b)
            If temp2 > 0# Then i = i + temp * (tpr - temp2) / (2# * temp2)
         End If
      Else
         critbetabinomial = i
         Exit Function
      End If
   Wend
End Function

Private Function critcompbetabinomial(ByVal a As Double, ByVal b As Double, ByVal ss As Double, ByVal cprob As Double) As Double
'//i such that 1-Pr(betabinomial(i,ss,a,b)) > cprob and  1-Pr(betabinomial(i-1,ss,a,b)) <= cprob
   If (cprob > 0.5) Then
      critcompbetabinomial = critbetabinomial(a, b, ss, 1# - cprob)
      Exit Function
   End If
   Dim pr As Double, tpr As Double
   Dim i As Double, temp As Double, temp2 As Double, oneMinusP As Double
   If (a + b < 1#) Then
      i = invcompbeta(a, b, cprob, oneMinusP) * ss
   Else
      pr = ss / (ss + a + b - 1#)
      i = invcompbeta(a * pr, b * pr, cprob, oneMinusP) * ss
   End If
   While (True)
      If (i < 0#) Then
         i = 0#
      ElseIf (i > ss) Then
         i = ss
      End If
      i = Int(i + 0.5)
      If (i >= max_discrete) Then
         critcompbetabinomial = i
         Exit Function
      End If
      pr = comp_cdf_BetaBinomial(i, ss, a, b)
      tpr = 0#
      If (pr >= cprob * (1 + cfSmall)) Then
         i = i + 1#
         tpr = pmf_BetaBinomial(i, ss, a, b)
         If (pr < (1.00001) * tpr) Then
            Do While (tpr > cprob)
               i = i + 1#
               temp = ss + b - i
               If temp = 0# Then Exit Do
               tpr = tpr * ((a + i - 1#) * (ss - i + 1#)) / (i * temp)
            Loop
         Else
            pr = pr - tpr
            If (pr <= cprob) Then
               critcompbetabinomial = i
               Exit Function
            End If
            temp = (pr - cprob) / tpr
            If (temp > 10#) Then
               temp = Int(temp + 0.5)
               i = i + temp
               temp2 = pmf_BetaBinomial(i, ss, a, b)
               i = i + temp * (tpr - temp2) / (2# * temp2)
            Else
               i = i + 1#
               tpr = tpr * ((a + i - 1#) * (ss - i + 1#)) / (i * (ss + b - i))
               pr = pr - tpr
               If (pr <= cprob) Then
                  critcompbetabinomial = i
                  Exit Function
               End If
               temp2 = (pr - cprob) / tpr
               If (temp2 < temp - 0.9) Then
                  While (pr > cprob)
                     i = i + 1#
                     tpr = tpr * ((a + i - 1#) * (ss - i + 1#)) / (i * (ss + b - i))
                     pr = pr - tpr
                  Wend
                  critcompbetabinomial = i
                  Exit Function
               Else
                  temp = Int(Log(cprob / pr) / Log(((a + i - 1#) * (ss - i + 1#)) / (i * (ss + b - i))) + 0.5)
                  i = i + temp
                  temp2 = pmf_BetaBinomial(i, ss, a, b)
                  If (temp2 > nearly_zero) Then
                     temp = Log((cprob / pr) * (tpr / temp2)) / Log(((a + i - 1#) * (ss - i + 1#)) / (i * (ss + b - i)))
                     i = i + temp
                  End If
               End If
            End If
         End If
      ElseIf ((1# + cfSmall) * pr < cprob) Then
         While ((tpr < cSmall) And (pr <= cprob))
            tpr = pmf_BetaBinomial(i, ss, a, b)
            pr = pr + tpr
            i = i - 1#
         Wend
         temp = (cprob - pr) / tpr
         If temp <= 0# Then
            critcompbetabinomial = i + 1#
            Exit Function
         ElseIf temp < 100# Or i < 1000# Then
            While (pr <= cprob)
               tpr = tpr * ((i + 1#) * (ss + b - i - 1#)) / ((a + i) * (ss - i))
               pr = pr + tpr
               i = i - 1#
            Wend
            critcompbetabinomial = i + 1#
            Exit Function
         ElseIf temp > i Then
            i = i / 10#
         Else
            i = Int(i - temp)
            temp2 = pmf_BetaNegativeBinomial(i, ss, a, b)
            If temp2 > 0# Then i = i - temp * (tpr - temp2) / (2# * temp2)
         End If
      Else
         critcompbetabinomial = i + 1#
         Exit Function
      End If
   Wend
End Function

Public Function crit_BetaBinomial(ByVal sample_size As Double, ByVal beta_shape1 As Double, ByVal beta_shape2 As Double, ByVal crit_prob As Double) As Double
   sample_size = AlterForIntegralChecks_Others(sample_size)
   If (beta_shape1 <= 0# Or beta_shape2 <= 0# Or sample_size < 0#) Then
      crit_BetaBinomial = [#VALUE!]
   ElseIf (crit_prob < 0# Or crit_prob > 1#) Then
      crit_BetaBinomial = [#VALUE!]
   ElseIf (crit_prob = 0#) Then
      crit_BetaBinomial = [#VALUE!]
   ElseIf (sample_size = 0# Or crit_prob = 1#) Then
      crit_BetaBinomial = sample_size
   Else
      Dim i As Double, pr As Double
      i = critbetabinomial(beta_shape1, beta_shape2, sample_size, crit_prob)
      crit_BetaBinomial = i
      pr = cdf_BetaBinomial(i, sample_size, beta_shape1, beta_shape2)
      If (pr = crit_prob) Then
      ElseIf (pr > crit_prob) Then
         i = i - 1#
         pr = cdf_BetaBinomial(i, sample_size, beta_shape1, beta_shape2)
         If (pr >= crit_prob) Then
            crit_BetaBinomial = i
         End If
      Else
         crit_BetaBinomial = i + 1#
      End If
   End If
   crit_BetaBinomial = GetRidOfMinusZeroes(crit_BetaBinomial)
End Function

Public Function comp_crit_BetaBinomial(ByVal sample_size As Double, ByVal beta_shape1 As Double, ByVal beta_shape2 As Double, ByVal crit_prob As Double) As Double
   sample_size = AlterForIntegralChecks_Others(sample_size)
   If (beta_shape1 <= 0# Or beta_shape2 <= 0# Or sample_size < 0#) Then
      comp_crit_BetaBinomial = [#VALUE!]
   ElseIf (crit_prob < 0# Or crit_prob > 1#) Then
      comp_crit_BetaBinomial = [#VALUE!]
   ElseIf (crit_prob = 1#) Then
      comp_crit_BetaBinomial = 0#
   ElseIf (sample_size = 0# Or crit_prob = 0#) Then
      comp_crit_BetaBinomial = sample_size
   Else
      Dim i As Double, pr As Double
      i = critcompbetabinomial(beta_shape1, beta_shape2, sample_size, crit_prob)
      comp_crit_BetaBinomial = i
      pr = comp_cdf_BetaBinomial(i, sample_size, beta_shape1, beta_shape2)
      If (pr = crit_prob) Then
      ElseIf (pr < crit_prob) Then
         i = i - 1#
         pr = comp_cdf_BetaBinomial(i, sample_size, beta_shape1, beta_shape2)
         If (pr <= crit_prob) Then
            comp_crit_BetaBinomial = i
         End If
      Else
         comp_crit_BetaBinomial = i + 1#
      End If
   End If
   comp_crit_BetaBinomial = GetRidOfMinusZeroes(comp_crit_BetaBinomial)
End Function

Public Function pdf_normal_os(ByVal x As Double, Optional ByVal n As Double = 1, Optional ByVal r As Double = -1) As Double
 ' pdf for rth smallest (r>0) or -rth largest (r<0) of the n order statistics from a sample of iid N(0,1) variables
 ' based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    If n < 1 Or Abs(r) > n Or r = 0 Then pdf_normal_os = [#VALUE!]: Exit Function
    Dim n1 As Double: n1 = n + 1
    If r > 0 Then r = r - n1
    If x <= 0 Then
        pdf_normal_os = pdf_beta(cnormal(x), n1 + r, -r) * pdf_normal(x)
    Else
        pdf_normal_os = pdf_beta(cnormal(-x), -r, n1 + r) * pdf_normal(-x)
    End If
End Function
 
Public Function cdf_normal_os(ByVal x As Double, Optional ByVal n As Double = 1, Optional ByVal r As Double = -1) As Double
 ' cdf for rth smallest (r>0) or -rth largest (r<0) of the n order statistics from a sample of iid N(0,1) variables
 ' based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    If n < 1 Or Abs(r) > n Or r = 0 Then cdf_normal_os = [#VALUE!]: Exit Function
    Dim n1 As Double: n1 = n + 1
    If r > 0 Then r = r - n1
    If x <= 0 Then
        cdf_normal_os = cdf_beta(cnormal(x), n1 + r, -r)
    Else
        cdf_normal_os = comp_cdf_beta(cnormal(-x), -r, n1 + r)
    End If
End Function
 
Public Function comp_cdf_normal_os(ByVal x As Double, Optional ByVal n As Double = 1, Optional ByVal r As Double = -1) As Double
 ' 1-cdf for rth smallest (r>0) or -rth largest (r<0) of the n order statistics from a sample of iid N(0,1) variables
 ' based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    If n < 1 Or Abs(r) > n Or r = 0 Then comp_cdf_normal_os = [#VALUE!]: Exit Function
    Dim n1 As Double: n1 = n + 1
    If r > 0 Then r = r - n1
    If x <= 0 Then
        comp_cdf_normal_os = comp_cdf_beta(cnormal(x), n1 + r, -r)
    Else
        comp_cdf_normal_os = cdf_beta(cnormal(-x), -r, n1 + r)
    End If
End Function
 
Public Function inv_normal_os(ByVal p As Double, Optional ByVal n As Double = 1, Optional ByVal r As Double = -1) As Double
 ' inverse of cdf_normal_os
 ' based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
 ' accuracy for median of extreme order statistic is limited by accuracy of IEEE double precision representation of n >> 10^15, not by this routine
    Dim oneMinusxp As Double
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    If n < 1 Or Abs(r) > n Or r = 0 Then inv_normal_os = [#VALUE!]: Exit Function
    Dim xp As Double, n1 As Double: n1 = n + 1
    If r > 0 Then r = r - n1
    xp = invbeta(n1 + r, -r, p, oneMinusxp)
    If Abs(xp - 0.5) < 0.00000000000001 And xp <> 0.5 Then If cdf_beta(0.5, n1 + r, -r) = p Then inv_normal_os = 0: Exit Function
    If xp <= 0.5 Then
        inv_normal_os = inv_normal(xp)
    Else
        inv_normal_os = -inv_normal(oneMinusxp)
    End If
End Function
 
Public Function comp_inv_normal_os(ByVal p As Double, Optional ByVal n As Double = 1, Optional ByVal r As Double = -1) As Double
 ' inverse of comp_cdf_normal_os
 ' based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
    Dim oneMinusxp As Double
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    If n < 1 Or Abs(r) > n Or r = 0 Then comp_inv_normal_os = [#VALUE!]: Exit Function
    Dim xp As Double, n1 As Double: n1 = n + 1
    If r > 0 Then r = r - n1
    xp = invcompbeta(n1 + r, -r, p, oneMinusxp)
    If Abs(xp - 0.5) < 0.00000000000001 And xp <> 0.5 Then If comp_cdf_beta(0.5, n1 + r, -r) = p Then comp_inv_normal_os = 0: Exit Function
    If xp <= 0.5 Then
        comp_inv_normal_os = inv_normal(xp)
    Else
        comp_inv_normal_os = -inv_normal(oneMinusxp)
    End If
End Function

Public Function pdf_gamma_os(ByVal x As Double, ByVal shape_param As Double, Optional ByVal n As Double = 1, Optional ByVal r As Double = -1, Optional ByVal scale_param As Double = 1, Optional ByVal nc_param As Double = 0) As Double
 ' pdf for rth smallest (r>0) or -rth largest (r<0) of the n order statistics from a sample of iid gamma(a,1) variables
 ' based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    If n < 1 Or Abs(r) > n Or r = 0 Then pdf_gamma_os = [#VALUE!]: Exit Function
    Dim p As Double, n1 As Double: n1 = n + 1
    If r > 0 Then r = r - n1
    p = cdf_gamma_nc(x / scale_param, shape_param, nc_param)
    If p <= 0.5 Then        ' avoid truncation error by working with p <= 0.5
        pdf_gamma_os = pdf_beta(p, n1 + r, -r) * pdf_gamma_nc(x / scale_param, shape_param, nc_param) / scale_param
    Else
        pdf_gamma_os = pdf_beta(comp_cdf_gamma_nc(x / scale_param, shape_param, nc_param), -r, n1 + r) * pdf_gamma_nc(x / scale_param, shape_param, nc_param) / scale_param
    End If
End Function

Public Function cdf_gamma_os(ByVal x As Double, ByVal shape_param As Double, Optional ByVal n As Double = 1, Optional ByVal r As Double = -1, Optional ByVal scale_param As Double = 1, Optional ByVal nc_param As Double = 0) As Double
 ' cdf for rth smallest (r>0) or -rth largest (r<0) of the n order statistics from a sample of iid gamma(a,1) variables
 ' based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    If n < 1 Or Abs(r) > n Or r = 0 Then cdf_gamma_os = [#VALUE!]: Exit Function
    Dim p As Double, n1 As Double: n1 = n + 1
    If r > 0 Then r = r - n1
    p = cdf_gamma_nc(x / scale_param, shape_param, nc_param)
    If p <= 0.5 Then        ' avoid truncation error by working with p <= 0.5
        cdf_gamma_os = cdf_beta(p, n1 + r, -r)
    Else
        cdf_gamma_os = comp_cdf_beta(comp_cdf_gamma_nc(x / scale_param, shape_param, nc_param), -r, n1 + r)
    End If
End Function

Public Function comp_cdf_gamma_os(ByVal x As Double, ByVal shape_param As Double, Optional ByVal n As Double = 1, Optional ByVal r As Double = -1, Optional ByVal scale_param As Double = 1, Optional ByVal nc_param As Double = 0) As Double
 ' 1-cdf for rth smallest (r>0) or -rth largest (r<0) of the n order statistics from a sample of iid gamma(a,1) variables
 ' based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    If n < 1 Or Abs(r) > n Or r = 0 Then comp_cdf_gamma_os = [#VALUE!]: Exit Function
    Dim p As Double, n1 As Double: n1 = n + 1
    If r > 0 Then r = r - n1
    p = cdf_gamma_nc(x / scale_param, shape_param, nc_param)
    If p <= 0.5 Then        ' avoid truncation error by working with p <= 0.5
        comp_cdf_gamma_os = comp_cdf_beta(p, n1 + r, -r)
    Else
        comp_cdf_gamma_os = cdf_beta(comp_cdf_gamma_nc(x / scale_param, shape_param, nc_param), -r, n1 + r)
    End If
End Function

Public Function inv_gamma_os(ByVal p As Double, ByVal shape_param As Double, Optional ByVal n As Double = 1, Optional ByVal r As Double = -1, Optional ByVal scale_param As Double = 1, Optional ByVal nc_param As Double = 0) As Double
 ' inverse of cdf_gamma_os
 ' based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
 ' accuracy for median of extreme order statistic is limited by accuracy of IEEE double precision representation of n >> 10^15, not by this routine
    Dim oneMinusxp As Double
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    If n < 1 Or Abs(r) > n Or r = 0 Then inv_gamma_os = [#VALUE!]: Exit Function
    Dim xp As Double, n1 As Double: n1 = n + 1
    If r > 0 Then r = r - n1
    xp = invbeta(n1 + r, -r, p, oneMinusxp)
    If xp <= 0.5 Then       ' avoid truncation error by working with xp <= 0.5
        inv_gamma_os = inv_gamma_nc(xp, shape_param, nc_param) * scale_param
    Else
        inv_gamma_os = comp_inv_gamma_nc(oneMinusxp, shape_param, nc_param) * scale_param
    End If
End Function

Public Function comp_inv_gamma_os(ByVal p As Double, ByVal shape_param As Double, Optional ByVal n As Double = 1, Optional ByVal r As Double = -1, Optional ByVal scale_param As Double = 1, Optional ByVal nc_param As Double = 0) As Double
 ' inverse of comp_cdf_gamma_os
 ' based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
    Dim oneMinusxp As Double
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    If n < 1 Or Abs(r) > n Or r = 0 Then comp_inv_gamma_os = [#VALUE!]: Exit Function
    Dim xp As Double, n1 As Double: n1 = n + 1
    If r > 0 Then r = r - n1
    xp = invcompbeta(n1 + r, -r, p, oneMinusxp)
    If xp <= 0.5 Then       ' avoid truncation error by working with xp <= 0.5
        comp_inv_gamma_os = inv_gamma_nc(xp, shape_param, nc_param) * scale_param
    Else
        comp_inv_gamma_os = comp_inv_gamma_nc(oneMinusxp, shape_param, nc_param) * scale_param
    End If
End Function

Public Function pdf_chi2_os(ByVal x As Double, ByVal df As Double, Optional ByVal n As Double = 1, Optional ByVal r As Double = -1, Optional ByVal nc_param As Double = 0) As Double
 ' pdf for rth smallest (r>0) or -rth largest (r<0) of the n order statistics from a sample of iid chi2(df) variables
 ' based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    If n < 1 Or Abs(r) > n Or r = 0 Then pdf_chi2_os = [#VALUE!]: Exit Function
   df = AlterForIntegralChecks_df(df)
    Dim p As Double, n1 As Double: n1 = n + 1
    If r > 0 Then r = r - n1
    p = cdf_Chi2_nc(x, df, nc_param)
    If p <= 0.5 Then        ' avoid truncation error by working with p <= 0.5
        pdf_chi2_os = pdf_beta(p, n1 + r, -r) * pdf_Chi2_nc(x, df, nc_param)
    Else
        pdf_chi2_os = pdf_beta(comp_cdf_Chi2_nc(x, df, nc_param), -r, n1 + r) * pdf_Chi2_nc(x, df, nc_param)
    End If
End Function

Public Function cdf_chi2_os(ByVal x As Double, ByVal df As Double, Optional ByVal n As Double = 1, Optional ByVal r As Double = -1, Optional ByVal nc_param As Double = 0) As Double
 ' cdf for rth smallest (r>0) or -rth largest (r<0) of the n order statistics from a sample of iid chi2(df) variables
 ' based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    If n < 1 Or Abs(r) > n Or r = 0 Then cdf_chi2_os = [#VALUE!]: Exit Function
   df = AlterForIntegralChecks_df(df)
    Dim p As Double, n1 As Double: n1 = n + 1
    If r > 0 Then r = r - n1
    p = cdf_Chi2_nc(x, df, nc_param)
    If p <= 0.5 Then        ' avoid truncation error by working with p <= 0.5
        cdf_chi2_os = cdf_beta(p, n1 + r, -r)
    Else
        cdf_chi2_os = comp_cdf_beta(comp_cdf_Chi2_nc(x, df, nc_param), -r, n1 + r)
    End If
End Function

Public Function comp_cdf_chi2_os(ByVal x As Double, ByVal df As Double, Optional ByVal n As Double = 1, Optional ByVal r As Double = -1, Optional ByVal nc_param As Double = 0) As Double
 ' 1-cdf for rth smallest (r>0) or -rth largest (r<0) of the n order statistics from a sample of iid chi2(df) variables
 ' based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    If n < 1 Or Abs(r) > n Or r = 0 Then comp_cdf_chi2_os = [#VALUE!]: Exit Function
   df = AlterForIntegralChecks_df(df)
    Dim p As Double, n1 As Double: n1 = n + 1
    If r > 0 Then r = r - n1
    p = cdf_Chi2_nc(x, df, nc_param)
    If p <= 0.5 Then        ' avoid truncation error by working with p <= 0.5
        comp_cdf_chi2_os = comp_cdf_beta(p, n1 + r, -r)
    Else
        comp_cdf_chi2_os = cdf_beta(comp_cdf_Chi2_nc(x, df, nc_param), -r, n1 + r)
    End If
End Function

Public Function inv_chi2_os(ByVal p As Double, ByVal df As Double, Optional ByVal n As Double = 1, Optional ByVal r As Double = -1, Optional ByVal nc_param As Double = 0) As Double
 ' inverse of cdf_chi2_os
 ' based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
 ' accuracy for median of extreme order statistic is limited by accuracy of IEEE double precision representation of n >> 10^15, not by this routine
    Dim oneMinusxp As Double
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    If n < 1 Or Abs(r) > n Or r = 0 Then inv_chi2_os = [#VALUE!]: Exit Function
   df = AlterForIntegralChecks_df(df)
    Dim xp As Double, n1 As Double: n1 = n + 1
    If r > 0 Then r = r - n1
    xp = invbeta(n1 + r, -r, p, oneMinusxp)
    If xp <= 0.5 Then       ' avoid truncation error by working with xp <= 0.5
        inv_chi2_os = inv_Chi2_nc(xp, df, nc_param)
    Else
        inv_chi2_os = comp_inv_Chi2_nc(oneMinusxp, df, nc_param)
    End If
End Function

Public Function comp_inv_chi2_os(ByVal p As Double, ByVal df As Double, Optional ByVal n As Double = 1, Optional ByVal r As Double = -1, Optional ByVal nc_param As Double = 0) As Double
 ' inverse of comp_cdf_chi2_os
 ' based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
    Dim oneMinusxp As Double
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    If n < 1 Or Abs(r) > n Or r = 0 Then comp_inv_chi2_os = [#VALUE!]: Exit Function
   df = AlterForIntegralChecks_df(df)
    Dim xp As Double, n1 As Double: n1 = n + 1
    If r > 0 Then r = r - n1
    xp = invcompbeta(n1 + r, -r, p, oneMinusxp)
    If xp <= 0.5 Then       ' avoid truncation error by working with xp <= 0.5
        comp_inv_chi2_os = inv_Chi2_nc(xp, df, nc_param)
    Else
        comp_inv_chi2_os = comp_inv_Chi2_nc(oneMinusxp, df, nc_param)
    End If
End Function

Public Function pdf_F_os(ByVal x As Double, ByVal df1 As Double, ByVal df2 As Double, Optional ByVal n As Double = 1, Optional ByVal r As Double = -1, Optional ByVal nc_param As Double = 0) As Double
 ' pdf for rth smallest (r>0) or -rth largest (r<0) of the n order statistics from a sample of iid gamma(a,1) variables
 ' based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    If n < 1 Or Abs(r) > n Or r = 0 Then pdf_F_os = [#VALUE!]: Exit Function
   df1 = AlterForIntegralChecks_df(df1)
   df2 = AlterForIntegralChecks_df(df2)
    Dim p As Double, n1 As Double: n1 = n + 1
    If r > 0 Then r = r - n1
    p = cdf_fdist_nc(x, df1, df2, nc_param)
    If p <= 0.5 Then        ' avoid truncation error by working with p <= 0.5
        pdf_F_os = pdf_beta(p, n1 + r, -r) * pdf_fdist_nc(x, df1, df2, nc_param)
    Else
        pdf_F_os = pdf_beta(comp_cdf_fdist_nc(x, df1, df2, nc_param), -r, n1 + r) * pdf_fdist_nc(x, df1, df2, nc_param)
    End If
End Function

Public Function cdf_F_os(ByVal x As Double, ByVal df1 As Double, ByVal df2 As Double, Optional ByVal n As Double = 1, Optional ByVal r As Double = -1, Optional ByVal nc_param As Double = 0) As Double
 ' cdf for rth smallest (r>0) or -rth largest (r<0) of the n order statistics from a sample of iid gamma(a,1) variables
 ' based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    If n < 1 Or Abs(r) > n Or r = 0 Then cdf_F_os = [#VALUE!]: Exit Function
   df1 = AlterForIntegralChecks_df(df1)
   df2 = AlterForIntegralChecks_df(df2)
    Dim p As Double, n1 As Double: n1 = n + 1
    If r > 0 Then r = r - n1
    p = cdf_fdist_nc(x, df1, df2, nc_param)
    If p <= 0.5 Then        ' avoid truncation error by working with p <= 0.5
        cdf_F_os = cdf_beta(p, n1 + r, -r)
    Else
        cdf_F_os = comp_cdf_beta(comp_cdf_fdist_nc(x, df1, df2, nc_param), -r, n1 + r)
    End If
End Function

Public Function comp_cdf_F_os(ByVal x As Double, ByVal df1 As Double, ByVal df2 As Double, Optional ByVal n As Double = 1, Optional ByVal r As Double = -1, Optional ByVal nc_param As Double = 0) As Double
 ' 1-cdf for rth smallest (r>0) or -rth largest (r<0) of the n order statistics from a sample of iid gamma(a,1) variables
 ' based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    If n < 1 Or Abs(r) > n Or r = 0 Then comp_cdf_F_os = [#VALUE!]: Exit Function
   df1 = AlterForIntegralChecks_df(df1)
   df2 = AlterForIntegralChecks_df(df2)
    Dim p As Double, n1 As Double: n1 = n + 1
    If r > 0 Then r = r - n1
    p = cdf_fdist_nc(x, df1, df2, nc_param)
    If p <= 0.5 Then        ' avoid truncation error by working with p <= 0.5
        comp_cdf_F_os = comp_cdf_beta(p, n1 + r, -r)
    Else
        comp_cdf_F_os = cdf_beta(comp_cdf_fdist_nc(x, df1, df2, nc_param), -r, n1 + r)
    End If
End Function

Public Function inv_F_os(ByVal p As Double, ByVal df1 As Double, ByVal df2 As Double, Optional ByVal n As Double = 1, Optional ByVal r As Double = -1, Optional ByVal nc_param As Double = 0) As Double
 ' inverse of cdf_F_os
 ' based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
 ' accuracy for median of extreme order statistic is limited by accuracy of IEEE double precision representation of n >> 10^15, not by this routine
    Dim oneMinusxp As Double
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    If n < 1 Or Abs(r) > n Or r = 0 Then inv_F_os = [#VALUE!]: Exit Function
   df1 = AlterForIntegralChecks_df(df1)
   df2 = AlterForIntegralChecks_df(df2)
    Dim xp As Double, n1 As Double: n1 = n + 1
    If r > 0 Then r = r - n1
    xp = invbeta(n1 + r, -r, p, oneMinusxp)
    If xp <= 0.5 Then       ' avoid truncation error by working with xp <= 0.5
        inv_F_os = inv_fdist_nc(xp, df1, df2, nc_param)
    Else
        inv_F_os = comp_inv_fdist_nc(oneMinusxp, df1, df2, nc_param)
    End If
End Function

Public Function comp_inv_F_os(ByVal p As Double, ByVal df1 As Double, ByVal df2 As Double, Optional ByVal n As Double = 1, Optional ByVal r As Double = -1, Optional ByVal nc_param As Double = 0) As Double
 ' inverse of comp_cdf_F_os
 ' based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
    Dim oneMinusxp As Double
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    If n < 1 Or Abs(r) > n Or r = 0 Then comp_inv_F_os = [#VALUE!]: Exit Function
   df1 = AlterForIntegralChecks_df(df1)
   df2 = AlterForIntegralChecks_df(df2)
    Dim xp As Double, n1 As Double: n1 = n + 1
    If r > 0 Then r = r - n1
    xp = invcompbeta(n1 + r, -r, p, oneMinusxp)
    If xp <= 0.5 Then       ' avoid truncation error by working with xp <= 0.5
        comp_inv_F_os = inv_fdist_nc(xp, df1, df2, nc_param)
    Else
        comp_inv_F_os = comp_inv_fdist_nc(oneMinusxp, df1, df2, nc_param)
    End If
End Function

Public Function pdf_beta_os(ByVal x As Double, ByVal shape_param1 As Double, ByVal shape_param2 As Double, Optional ByVal n As Double = 1, Optional ByVal r As Double = -1, Optional ByVal nc_param As Double = 0) As Double
 ' pdf for rth smallest (r>0) or -rth largest (r<0) of the n order statistics from a sample of iid gamma(a,1) variables
 ' based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    If n < 1 Or Abs(r) > n Or r = 0 Then pdf_beta_os = [#VALUE!]: Exit Function
    Dim p As Double, n1 As Double: n1 = n + 1
    If r > 0 Then r = r - n1
    p = cdf_beta_nc(x, shape_param1, shape_param2, nc_param)
    If p <= 0.5 Then        ' avoid truncation error by working with p <= 0.5
        pdf_beta_os = pdf_beta(p, n1 + r, -r) * pdf_beta_nc(x, shape_param1, shape_param2, nc_param)
    Else
        pdf_beta_os = pdf_beta(comp_cdf_beta_nc(x, shape_param1, shape_param2, nc_param), -r, n1 + r) * pdf_beta_nc(x, shape_param1, shape_param2, nc_param)
    End If
End Function

Public Function cdf_beta_os(ByVal x As Double, ByVal shape_param1 As Double, ByVal shape_param2 As Double, Optional ByVal n As Double = 1, Optional ByVal r As Double = -1, Optional ByVal nc_param As Double = 0) As Double
 ' cdf for rth smallest (r>0) or -rth largest (r<0) of the n order statistics from a sample of iid gamma(a,1) variables
 ' based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    If n < 1 Or Abs(r) > n Or r = 0 Then cdf_beta_os = [#VALUE!]: Exit Function
    Dim p As Double, n1 As Double: n1 = n + 1
    If r > 0 Then r = r - n1
    p = cdf_beta_nc(x, shape_param1, shape_param2, nc_param)
    If p <= 0.5 Then        ' avoid truncation error by working with p <= 0.5
        cdf_beta_os = cdf_beta(p, n1 + r, -r)
    Else
        cdf_beta_os = comp_cdf_beta(comp_cdf_beta_nc(x, shape_param1, shape_param2, nc_param), -r, n1 + r)
    End If
End Function

Public Function comp_cdf_beta_os(ByVal x As Double, ByVal shape_param1 As Double, ByVal shape_param2 As Double, Optional ByVal n As Double = 1, Optional ByVal r As Double = -1, Optional ByVal nc_param As Double = 0) As Double
 ' 1-cdf for rth smallest (r>0) or -rth largest (r<0) of the n order statistics from a sample of iid gamma(a,1) variables
 ' based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    If n < 1 Or Abs(r) > n Or r = 0 Then comp_cdf_beta_os = [#VALUE!]: Exit Function
    Dim p As Double, n1 As Double: n1 = n + 1
    If r > 0 Then r = r - n1
    p = cdf_beta_nc(x, shape_param1, shape_param2, nc_param)
    If p <= 0.5 Then        ' avoid truncation error by working with p <= 0.5
        comp_cdf_beta_os = comp_cdf_beta(p, n1 + r, -r)
    Else
        comp_cdf_beta_os = cdf_beta(comp_cdf_beta_nc(x, shape_param1, shape_param2, nc_param), -r, n1 + r)
    End If
End Function

Public Function inv_beta_os(ByVal p As Double, ByVal shape_param1 As Double, ByVal shape_param2 As Double, Optional ByVal n As Double = 1, Optional ByVal r As Double = -1, Optional ByVal nc_param As Double = 0) As Double
 ' inverse of cdf_beta_os
 ' based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
 ' accuracy for median of extreme order statistic is limited by accuracy of IEEE double precision representation of n >> 10^15, not by this routine
    Dim oneMinusxp As Double
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    If n < 1 Or Abs(r) > n Or r = 0 Then inv_beta_os = [#VALUE!]: Exit Function
    Dim xp As Double, n1 As Double: n1 = n + 1
    If r > 0 Then r = r - n1
    xp = invbeta(n1 + r, -r, p, oneMinusxp)
    If xp <= 0.5 Then       ' avoid truncation error by working with xp <= 0.5
        inv_beta_os = inv_beta_nc(xp, shape_param1, shape_param2, nc_param)
    Else
        inv_beta_os = comp_inv_beta_nc(oneMinusxp, shape_param1, shape_param2, nc_param)
    End If
End Function

Public Function comp_inv_beta_os(ByVal p As Double, ByVal shape_param1 As Double, ByVal shape_param2 As Double, Optional ByVal n As Double = 1, Optional ByVal r As Double = -1, Optional ByVal nc_param As Double = 0) As Double
 ' inverse of comp_cdf_beta_os
 ' based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
    Dim oneMinusxp As Double
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    If n < 1 Or Abs(r) > n Or r = 0 Then comp_inv_beta_os = [#VALUE!]: Exit Function
    Dim xp As Double, n1 As Double: n1 = n + 1
    If r > 0 Then r = r - n1
    xp = invcompbeta(n1 + r, -r, p, oneMinusxp)
    If xp <= 0.5 Then       ' avoid truncation error by working with xp <= 0.5
        comp_inv_beta_os = inv_beta_nc(xp, shape_param1, shape_param2, nc_param)
    Else
        comp_inv_beta_os = comp_inv_beta_nc(oneMinusxp, shape_param1, shape_param2, nc_param)
    End If
End Function

Public Function fet_pearson22(a As Double, b As Double, c As Double, d As Double) As Double
'The following is some VBA code for the two-sided 2x2 FET based on Pearson's
'Chi-Square statistic (i.e. includes all tables which give a value of the
'Chi-Square statistic which is greater than or equal to that of the table observed)
Dim det As Double, temp As Double, sample_size As Double, pop As Double
det = a * d - b * c
If det > 0 Then
    temp = a
    a = b
    b = temp
    temp = c
    c = d
    d = temp
    det = -det
End If
sample_size = a + b
temp = a + c
pop = sample_size + c + d
det = (2# * det + 1) / pop
If det < -1# Then
   fet_pearson22 = cdf_hypergeometric(a, sample_size, temp, pop) + comp_cdf_hypergeometric(a - det, sample_size, temp, pop)
Else
   fet_pearson22 = 1#
End If
End Function

Public Function chi_square_test(r As Range) As Double
Dim cs As Double, rs As Double
Dim rc As Long, cc As Long, i As Long, j As Long, k As Long
rc = r.Rows.count
cc = r.Columns.count
If rc < 2 Or cc < 2 Then
   chi_square_test = [#VALUE!]
   Exit Function
End If
ReDim os(1 To rc, 1 To cc) As Double, Es(0 To rc, 0 To cc) As Double
For i = 1 To rc
    For j = 1 To cc
        os(i, j) = r.Item(i, j)
    Next j
Next i
'Calculate row totals and check that all values are non-negative integers
cs = 0#
For i = 1 To rc
   rs = 0#
   For j = 1 To cc
      If os(i, j) < 0 Or Int(os(i, j)) <> os(i, j) Then
         chi_square_test = [#VALUE!]
         Exit Function
      End If
      rs = rs + os(i, j)
   Next j
   Es(i, 0) = rs
   cs = cs + rs
Next i
Es(0, 0) = cs
'Calculate column totals
For i = 1 To cc
   rs = 0#
   For j = 1 To rc
      rs = rs + os(j, i)
   Next j
   Es(0, i) = rs
Next i
'Calculate chi_square value
rs = 0#
For i = 1 To rc
   For j = 1 To cc
      Es(i, j) = Es(i, 0) * Es(0, j) / cs
      rs = rs + (os(i, j) - Es(i, j)) ^ 2 / Es(i, j)
   Next j
Next i
chi_square_test = comp_cdf_chi_sq(rs, (rc - 1) * (cc - 1))

End Function

Public Function nidf_fdist(ByVal x As Double, ByVal df1 As Double, ByVal df2 As Double) As Double
   If (df1 <= 0# Or df2 <= 0#) Then
      nidf_fdist = [#VALUE!]
   ElseIf (x <= 0#) Then
      nidf_fdist = 0#
   Else
      Dim p As Double, q As Double
      p = df1 * x
      q = df2 + p
      p = p / q
      q = df2 / q
      df2 = df2 / 2#
      df1 = df1 / 2#
      If (p < 0.5) Then
          nidf_fdist = beta(p, df1, df2)
      Else
          nidf_fdist = compbeta(q, df2, df1)
      End If
   End If
End Function

Public Function comp_nidf_fdist(ByVal x As Double, ByVal df1 As Double, ByVal df2 As Double) As Double
   If (df1 <= 0# Or df2 <= 0#) Then
      comp_nidf_fdist = [#VALUE!]
   ElseIf (x <= 0#) Then
      comp_nidf_fdist = 1#
   Else
      Dim p As Double, q As Double
      p = df1 * x
      q = df2 + p
      p = p / q
      q = df2 / q
      df2 = df2 / 2#
      df1 = df1 / 2#
      If (p < 0.5) Then
          comp_nidf_fdist = compbeta(p, df1, df2)
      Else
          comp_nidf_fdist = beta(q, df2, df1)
      End If
   End If
End Function

Public Function CBNB(ByVal i As Double, ByVal r As Double, ByVal beta_shape1 As Double, ByVal beta_shape2 As Double) As Double
   Dim j As Double
   j = Int(i)
   CBNB = 0#
   Do While j > -1
      CBNB = CBNB + pmf_BetaNegativeBinomial(j, r, beta_shape1, beta_shape2)
      j = j - 1
   Loop
End Function

Public Function CBNB1(ByVal i As Double, ByVal r As Double, ByVal beta_shape1 As Double, ByVal beta_shape2 As Double) As Double
   Dim j As Double, ss As Double, bs2 As Double, temp As Double, swap As Double
   On Error GoTo errorhandler
   j = Int(i)
   ss = Min(r, beta_shape2)
   bs2 = Max(r, beta_shape2)
   CBNB1 = 0#
   temp = 0#
   If beta_shape1 >= 1# Then
      beta_shape1 = beta_shape1 - 1#
      temp = hypergeometricTerm(j, bs2, ss, beta_shape1)
      CBNB1 = temp
      Do While (beta_shape1 >= 1#) And (temp > 1E-16 * CBNB1)
         bs2 = bs2 + 1#
         ss = ss + 1#
         temp = temp * (j * beta_shape1) / (bs2 * ss)
         j = j - 1#
         beta_shape1 = beta_shape1 - 1#
         CBNB1 = CBNB1 + temp
      Loop
      If beta_shape1 >= 1# Then Exit Function
      j = j - 1#
      bs2 = bs2 + 1#
      ss = ss + 1#
   End If
   If temp = 0# Then
      temp = PBB(j, ss, bs2, beta_shape1)
   Else
      temp = temp * ((j + 1) * beta_shape1) / (ss * (j + bs2))
   End If
   CBNB1 = CBNB1 + temp
   Do While j > 0
      temp = (j * (ss + beta_shape1)) * temp
      ss = ss + 1
      j = j - 1
      If ss > bs2 Then
         swap = ss
         ss = bs2
         bs2 = swap
      End If
      temp = temp / ((j + bs2) * ss)
      CBNB1 = CBNB1 + temp
      If temp < 1E-16 * CBNB1 Then Exit Do
   Loop
'Debug.Print j, ss, bs2, beta_shape1
   Exit Function
errorhandler: Debug.Print j, ss, bs2, beta_shape1
End Function

Public Function ccBNB(ByVal i As Double, ByVal r As Double, ByVal beta_shape1 As Double, ByVal beta_shape2 As Double) As Double
   Dim ha1 As Double, hprob As Double, hswap As Boolean
   Dim mrb2 As Double, other As Double, temp As Double, ctemp As Double, mnib1 As Double, mxib1 As Double, swap As Double, max_iterations As Double
   i = Int(i)
   mrb2 = Max(r, beta_shape2)
   other = Min(r, beta_shape2)
   If (other <= 0# Or beta_shape1 <= 0#) Then
      ccBNB = [#VALUE!]
   ElseIf i < 0# Then
      ccBNB = 1#
   ElseIf (mrb2 >= 100# Or (mrb2 >= 5# And other * mrb2 > (i + 0.5) * (beta_shape1 + 0.5))) Then
      ccBNB = CBNB0(mrb2 - 1#, i + 1#, other, beta_shape1, 0#)
   Else
      mxib1 = beta_shape1
      mnib1 = i + 1#
      temp = pmf_BetaNegativeBinomial(mnib1, r, mxib1, beta_shape2)
      ctemp = temp
      max_iterations = Max(60#, mrb2)
      Do While (mnib1 + mxib1 < max_iterations) And temp > 1E-16 * ctemp
         temp = temp * (mnib1 + r) * (mnib1 + beta_shape2) / (mnib1 + mxib1 + r + beta_shape2)
         mnib1 = mnib1 + 1#
         If mxib1 < mnib1 Then
            swap = mxib1
            mxib1 = mnib1
            mnib1 = swap
         End If
         temp = temp / mnib1
         ctemp = ctemp + temp
      Loop
      temp = hypergeometricTerm(r, mnib1, mxib1, beta_shape2) * (r * beta_shape2 * (mnib1 + r + mxib1 + beta_shape2)) / ((mnib1 + 1#) * (mxib1 + r) * (mxib1 + beta_shape2))
      ctemp = ctemp + temp
      mnib1 = mnib1 + 1#
      r = r - 1#
      beta_shape2 = beta_shape2 - 1#
      Do
         mnib1 = mnib1 + 1#
         mxib1 = mxib1 + 1#
         temp = temp * (r * beta_shape2) / (mnib1 * mxib1)
         r = r - 1#
         beta_shape2 = beta_shape2 - 1#
         ctemp = ctemp + temp
      Loop Until Abs(temp) <= 1E-16 * ctemp
      ccBNB = ctemp
   End If
   ccBNB = GetRidOfMinusZeroes(ccBNB)
End Function

Function ccBNB5(ByVal ilim As Double, ByVal rr As Double, ByVal a As Double, ByVal bb As Double) As Double
   Dim temp As Double, i As Double, r As Double, b As Double
   If rr > bb Then
      r = rr
      b = bb
   Else
      r = bb
      b = rr
   End If
   ccBNB5 = (a + 0.5) * log0(b * r / ((a + 1#) * (b + a + r + 1#))) - r * log0(b / (r + a + 1#)) - b * log0(r / (a + b + 1#))
   If r <= 0.001 Then
      temp = a + (b + r) * 0.5
      ccBNB5 = ccBNB5  - b * r * (logfbit2(temp) + (b ^ 2 + r ^ 2) * logfbit4(temp) / 24#)
   Else
      ccBNB5 = ccBNB5  + (lfbaccdif1(b, r + a) - lfbaccdif1(b, a))
   End If
   temp = 0#
   If ilim > 0# Then
      i = ilim
      Do While i > 1#
         i = i - 1#
         temp = (1# + temp) * (i + r) * (i + b) / ((i + r + a + b) * (i + 1#))
      Loop
      temp = (1# + temp) * Exp(ccBNB5) * a
   End If
   ccBNB5 = (r * b * (1# - temp) - expm1(ccBNB5) * a * (r + a + b)) / ((r + a) * (a + b))
End Function

Function fet_22(c As Long, colsum() As Double, rowsum As Double, pmf_Obs As Double, ByRef inumstart As Double, ByRef jnumstart As Double) As Double
'The following is some VBA code for the two-sided 2x2 FET based on Pearson's
'Chi-Square statistic (i.e. includes all tables which give a value of the
'Chi-Square statistic which is greater than or equal to that of the table observed)
Dim inum_min As Double, jnum_max As Double, pmf_table As Double, inum As Double, jnum As Double, mode As Double, pmrc As Double, pop As Double, d As Double, prob_d As Double, pmfh As Double, pmfh_save As Double, knum As Double, prob As Double
Dim i As Long, j As Long
Dim all_d_zero As Boolean
ReDim ml(1 To c) As Double

'c = High(colsum) 'But can't pass partial arrays in calls
If pmf_Obs >= 1# Then  'All tables have pmf <= 1
   fet_22 = 1#
   Exit Function
End If
pop = 0#
For i = 1 To c
   pop = pop + colsum(i)
Next i

If c = 2 Then
   pmrc = pop - rowsum - colsum(2)
   mode = Int((rowsum + 1#) * (colsum(2) + 1#) / (pop + 2#))
   inum_min = Max(0#, -pmrc)
   inum = Min(Max(inum_min, inumstart), mode)
   jnum_max = Min(rowsum, colsum(2))
   jnum = Max(Min(jnum_max, jnumstart), mode)
   pmf_table = pmf_hypergeometric(inum, rowsum, colsum(2), pop)
   Do While pmf_table = 0#
      inum = Int((inum + Max(2, mode)) * 0.5)
      pmf_table = pmf_hypergeometric(inum, rowsum, colsum(2), pop)
   Loop
   Do While (pmf_table > pmf_Obs) And (inum > inum_min)
      pmf_table = pmf_table * (inum * (pmrc + inum))
      inum = inum - 1#
      pmf_table = pmf_table / ((rowsum - inum) * (colsum(2) - inum))
   Loop
   Do While (pmf_table <= pmf_Obs) And (inum <= mode)
      pmf_table = pmf_table * ((rowsum - inum) * (colsum(2) - inum))
      inum = inum + 1#
      pmf_table = pmf_table / (inum * (pmrc + inum))
   Loop
   pmf_table = pmf_hypergeometric(jnum, rowsum, colsum(2), pop)
   Do While pmf_table = 0#
      jnum = Int((jnum + mode) * 0.5)
      pmf_table = pmf_hypergeometric(jnum, rowsum, colsum(2), pop)
   Loop
   Do While (pmf_table > pmf_Obs) And (jnum < jnum_max)
      pmf_table = pmf_table * ((rowsum - jnum) * (colsum(2) - jnum))
      jnum = jnum + 1#
      pmf_table = pmf_table / (jnum * (pmrc + jnum))
   Loop
   Do While (pmf_table < pmf_Obs) And (jnum >= mode)
      pmf_table = pmf_table * (jnum * (pmrc + jnum))
      jnum = jnum - 1#
      pmf_table = pmf_table / ((rowsum - jnum) * (colsum(2) - jnum))
   Loop
   inumstart = inum
   jnumstart = jnum
   If inum > jnum Then
      fet_22 = 1#
   Else
      fet_22 = cdf_hypergeometric(inum - 1, rowsum, colsum(2), pop) + comp_cdf_hypergeometric(jnum, rowsum, colsum(2), pop)
   End If
Else
'First guess at mode vector
   ml(1) = rowsum
   For i = 2 To c
      ml(i) = Int(rowsum * colsum(i) / pop + 0.5)
      ml(1) = ml(1) - ml(i)
   Next i
   
   Do 'Update guess at mode vector
      all_d_zero = True
      For i = 1 To c - 1
         For j = i + 1 To c
            d = ml(i) - Int((colsum(i) + 1#) * (ml(i) + ml(j) + 1#) / (colsum(i) + colsum(j) + 2#))
            If d <> 0# Then
               ml(i) = ml(i) - d
               ml(j) = ml(j) + d
               all_d_zero = False
            End If
         Next j
      Next i
   Loop Until all_d_zero
   knum = ml(c)
   pmfh = pmf_hypergeometric(knum, rowsum, colsum(c), pop)
   If pmfh = 0# Then  'Not entirely sure what we want here but not likely that many tables have pmf < 1e-4933 and if there are it will be vary slow!
      fet_22 = "Probability of table is 0"
      Exit Function
   End If
   pmfh_save = pmfh
   inum = inumstart
   jnum = jnumstart
   prob_d = fet_22(c - 1, colsum, rowsum - knum, pmf_Obs / pmfh, inum, jnum)
'Debug.Print knum, pmfh, prob_d
   inumstart = inum
   jnumstart = jnum
   If inum > jnum Then
      fet_22 = 1#
      Exit Function
   End If
   prob = pmfh * prob_d
   Do
      pmfh = pmfh * (knum * (pop - colsum(c) - rowsum + knum))
      knum = knum - 1#
      pmfh = pmfh / ((colsum(c) - knum) * (rowsum - knum))
      If pmfh = 0# Then Exit Do
      prob_d = fet_22(c - 1, colsum, rowsum - knum, pmf_Obs / pmfh, inum, jnum)
'Debug.Print knum, pmfh, prob_d
      If inum > jnum Then
         prob = prob + cdf_hypergeometric(knum, colsum(c), rowsum, pop)
         Exit Do
      End If
      prob = prob + pmfh * prob_d
   Loop

   pmfh = pmfh_save
   knum = ml(c)
   inum = inumstart
   jnum = jnumstart
   Do
      pmfh = pmfh * ((colsum(c) - knum) * (rowsum - knum))
      knum = knum + 1#
      pmfh = pmfh / (knum * (pop - colsum(c) - rowsum + knum))
      If pmfh = 0# Then Exit Do
      prob_d = fet_22(c - 1, colsum, rowsum - knum, pmf_Obs / pmfh, inum, jnum)
'Debug.Print knum, pmfh, prob_d
      If inum > jnum Then
         prob = prob + comp_cdf_hypergeometric(knum - 1#, colsum(c), rowsum, pop)
         Exit Do
      End If
      prob = prob + pmfh * prob_d
   Loop
   fet_22 = prob
End If

End Function

Function old_fet_22(ByVal a As Double, ByVal b As Double, ByVal c As Double, ByVal d As Double) As Double
'The following is some VBA code for the two-sided 2x2 FET based on Pearson's
'Chi-Square statistic (i.e. includes all tables which give a value of the
'Chi-Square statistic which is greater than or equal to that of the table observed)
Dim det As Double, temp As Double, sample_size As Double, pop As Double, pmf_Obs As Double, pmf_table As Double, jnum As Double, mode As Double
det = GeneralabMinuscd(a, d, b, c)
If det > 0 Then
temp = a
a = b
b = temp
temp = c
c = d
d = temp
det = -det
End If
sample_size = a + b
temp = a + c
pop = sample_size + c + d
det = Int((2 * det + 1) / pop)
pmf_Obs = pmf_hypergeometric(a, sample_size, temp, pop) * (1.0000000000001)
If pmf_Obs = 0# Then
   old_fet_22 = 0#
   Exit Function
End If
mode = Int((sample_size + 1#) * (temp + 1#) / (pop + 2#))
jnum = a - det
pmf_table = pmf_hypergeometric(jnum, sample_size, temp, pop)
Do While pmf_table = 0#
   jnum = Int((jnum + mode) * 0.5)
   pmf_table = pmf_hypergeometric(jnum, sample_size, temp, pop)
Loop
If pmf_table > pmf_Obs Then
   Do While pmf_table >= pmf_Obs
      pmf_table = pmf_table * ((sample_size - jnum) * (temp - jnum))
      jnum = jnum + 1#
      pmf_table = pmf_table / (jnum * (pop - sample_size - temp + jnum))
   Loop
   jnum = jnum - 1#
Else
   Do While pmf_table <= pmf_Obs And jnum >= mode
      pmf_table = pmf_table * (jnum * (pop - sample_size - temp + jnum))
      jnum = jnum - 1#
      pmf_table = pmf_table / ((sample_size - jnum) * (temp - jnum))
   Loop
End If
If a > jnum Then
   old_fet_22 = 1#
Else
   old_fet_22 = cdf_hypergeometric(a, sample_size, temp, pop) + comp_cdf_hypergeometric(jnum, sample_size, temp, pop)
End If
End Function

Function fet_23(c As Long, ByRef colsum() As Double, ByVal rowsum As Double, ByVal pmf_Obs As Double, ByRef inumstart As Double, ByRef jnumstart As Double) As Double
Dim d As Double, cs As Double, colsum12 As Double, prob As Double, pmfh As Double, pmfh_save As Double, temp As Double, inum As Double, jnum As Double, knum As Double
Dim cdf As Double, ccdf As Double, pmf_table As Double, mode As Double, cdf_save As Double, ccdf_save As Double, col1mRowSum As Double, prob_d As Double, htTemp As Double
Dim pmf_table_inum As Double, pmf_table_jnum As Double, pmf_table_inum_save As Double, pmf_table_jnum_save As Double
Dim i As Long, j As Long, k As Long
Dim all_d_zero As Boolean
Dim ast As TAddStack

ReDim ml(1 To c) As Double

If pmf_Obs > 1# Then 'All tables have pmf <= 1
   fet_23 = 1#
   Exit Function
End If

cs = 0#
For i = 1 To c
   cs = cs + colsum(i)
Next i
'First guess at mode vector
ml(1) = rowsum
For i = 2 To c
   ml(i) = Int(rowsum * colsum(i) / cs + 0.5)
   ml(1) = ml(1) - ml(i)
Next i

Do 'Update guess at mode vector
   all_d_zero = True
   For i = 1 To c - 1
      For j = i + 1 To c
         d = ml(i) - Int((colsum(i) + 1#) * (ml(i) + ml(j) + 1#) / (colsum(i) + colsum(j) + 2#))
         If d <> 0# Then
            ml(i) = ml(i) - d
            ml(j) = ml(j) + d
            all_d_zero = False
         End If
      Next j
   Next i
Loop Until all_d_zero
knum = ml(c)
pmfh = pmf_hypergeometric(knum, rowsum, colsum(c), cs)
pmfh_save = pmfh

If c = 3 Then
   colsum12 = colsum(1) + colsum(2)
   col1mRowSum = colsum(1) - rowsum
   inum = Max(Max(0#, -(knum + col1mRowSum)), inumstart)
   jnum = Min(Min(rowsum - knum, colsum(2)), jnumstart)
   mode = Int((rowsum - knum + 1#) * (colsum(2) + 1#) / (colsum12 + 2#))
   If inum > mode Then
      inum = mode
   End If
   If jnum < mode Then
      jnum = mode
   End If
   pmf_table = pmf_hypergeometric(inum, rowsum - knum, colsum(2), colsum12)
   Do While pmf_table = 0#
      inum = Int((inum + Max(2, mode)) * 0.5)
      pmf_table = pmf_hypergeometric(inum, rowsum - knum, colsum(2), colsum12)
   Loop
   If pmf_table * pmfh <= pmf_Obs Then
      Do
         pmf_table = pmf_table * ((rowsum - knum - inum) * (colsum(2) - inum))
         inum = inum + 1#
         pmf_table = pmf_table / (inum * (col1mRowSum + knum + inum))
      Loop Until (pmf_table * pmfh > pmf_Obs) Or (inum > mode)
      pmf_table_inum = pmf_table
   Else
      Do
         pmf_table_inum = pmf_table
         pmf_table = pmf_table * (inum * (col1mRowSum + knum + inum))
         inum = inum - 1#
         pmf_table = pmf_table / ((rowsum - knum - inum) * (colsum(2) - inum))
      Loop Until pmf_table * pmfh <= pmf_Obs
      inum = inum + 1#
   End If
   pmf_table = pmf_hypergeometric(jnum, rowsum - knum, colsum(2), colsum12)
   Do While pmf_table = 0#
      jnum = Int((jnum + mode) * 0.5)
      pmf_table = pmf_hypergeometric(jnum, rowsum - knum, colsum(2), colsum12)
   Loop
   If pmf_table * pmfh > pmf_Obs Then
      Do
         pmf_table_jnum = pmf_table
         pmf_table = pmf_table * ((rowsum - knum - jnum) * (colsum(2) - jnum))
         jnum = jnum + 1#
         pmf_table = pmf_table / (jnum * (col1mRowSum + knum + jnum))
      Loop Until pmf_table * pmfh <= pmf_Obs
      jnum = jnum - 1#
   Else
      Do
         pmf_table = pmf_table * (jnum * (col1mRowSum + knum + jnum))
         jnum = jnum - 1#
         pmf_table = pmf_table / ((rowsum - knum - jnum) * (colsum(2) - jnum))
      Loop Until (pmf_table * pmfh > pmf_Obs) Or (jnum < mode)
      pmf_table_jnum = pmf_table
   End If
   inumstart = inum
   jnumstart = jnum
   If inum > jnum Then
      fet_23 = 1#
      Exit Function
   End If
   pmf_table_inum_save = pmf_table_inum
   pmf_table_jnum_save = pmf_table_jnum
   cdf = cdf_hypergeometric(inum - 1, rowsum - knum, colsum(2), colsum12)
   ccdf = comp_cdf_hypergeometric(jnum, rowsum - knum, colsum(2), colsum12)
   prob = 0#
   Call InitAddStack(ast)
   Call AddValueToStack(ast, pmfh * (cdf + ccdf))
'Debug.Print knum, pmfh * (cdf + ccdf)
   cdf_save = cdf
   ccdf_save = ccdf
   k = 0
   Do While knum >= 0
      k = k + 1
      If k = 100 Then
         knum = knum - 1#
         pmfh = pmf_hypergeometric(knum, rowsum, colsum(3), cs)
         k = 0
      Else
         pmfh = pmfh * (knum * (colsum12 - rowsum + knum))
         knum = knum - 1#
         pmfh = pmfh / ((rowsum - knum) * (colsum(3) - knum))
         'pmfh = pmf_hypergeometric(knum, rowsum, colsum(3), cs)
      End If
      If pmfh <= pmf_Obs Then Exit Do
      mode = Int((rowsum - knum + 1#) * (colsum(2) + 1#) / (colsum12 + 2#))
'If knum = 4294567294# Then
'   Debug.Print "Got here"
'End If
      inum = inum + 1#
      pmf_table = pmf_table_inum * ((rowsum - knum) * (colsum(2) - inum + 1#)) / (inum * (colsum12 - rowsum + knum + 1#))      'pmf_hypergeometric(inum, rowsum - knum, colsum(2), colsum12)
      temp = pmf_table_inum * (col1mRowSum + knum + inum) / (colsum12 - rowsum + knum + 1#)   'PBB(inum-1, colsum(2) - inum+1, rowsum - knum - inum+1, col1mRowSum + knum + inum)
      If pmf_table = 0# Then
         inum = inum - 1#
         pmf_table = temp * ((rowsum - knum)) / ((rowsum - knum - inum))  'pmf_hypergeometric(inum, rowsum - knum, colsum(2), colsum12)
         temp = temp - pmf_table
      End If
      Do While inum > mode
         pmf_table = pmf_table * (inum * (col1mRowSum + knum + inum))
         inum = inum - 1#
         pmf_table = pmf_table / ((rowsum - knum - inum) * (colsum(2) - inum))
         temp = temp - pmf_table
      Loop
      If pmf_table * pmfh <= pmf_Obs Then
         Do
            temp = temp + pmf_table
            pmf_table = pmf_table * ((rowsum - knum - inum) * (colsum(2) - inum))
            inum = inum + 1#
            pmf_table = pmf_table / (inum * (col1mRowSum + knum + inum))
         Loop Until (pmf_table * pmfh > pmf_Obs) Or (inum > mode)
         pmf_table_inum = pmf_table
      Else
         Do
            pmf_table_inum = pmf_table
            pmf_table = pmf_table * (inum * (col1mRowSum + knum + inum))
            inum = inum - 1#
            pmf_table = pmf_table / ((rowsum - knum - inum) * (colsum(2) - inum))
            If pmf_table = 0# Then
               cdf = 0#
               temp = 0#
               Exit Do
            End If
            If pmf_table * pmfh <= pmf_Obs Then Exit Do
            temp = temp - pmf_table
         Loop
         inum = inum + 1#
      End If
      If k = 50 Then pmf_table_inum = pmf_hypergeometric(inum, rowsum - knum, colsum(2), colsum12)
      cdf = cdf + temp
      pmf_table = pmf_table_jnum * ((rowsum - knum) * (col1mRowSum + knum + jnum + 1#)) / ((rowsum - knum - jnum) * (colsum12 - rowsum + knum + 1#)) 'pmf_hypergeometric(jnum, rowsum - knum, colsum(2), colsum12)
      temp = pmf_table_jnum * (colsum(2) - jnum) / (colsum12 - rowsum + knum + 1#)
      If pmf_table = 0# Then
         jnum = jnum + 1#
         pmf_table = temp * (rowsum - knum) / jnum 'pmf_hypergeometric(jnum, rowsum - knum, colsum(2), colsum12)
         temp = temp - pmf_table
      End If
      Do While jnum < mode
         pmf_table = pmf_table * ((rowsum - knum - jnum) * (colsum(2) - jnum))
         jnum = jnum + 1#
         pmf_table = pmf_table / (jnum * (col1mRowSum + knum + jnum))
         temp = temp - pmf_table
      Loop
      If pmf_table * pmfh <= pmf_Obs Then
         Do
            temp = temp + pmf_table
            pmf_table = pmf_table * (jnum * (col1mRowSum + knum + jnum))
            jnum = jnum - 1#
            pmf_table = pmf_table / ((rowsum - knum - jnum) * (colsum(2) - jnum))
         Loop Until (pmf_table * pmfh > pmf_Obs) Or (jnum < mode)
         pmf_table_jnum = pmf_table
      Else
         Do
            pmf_table_jnum = pmf_table
            pmf_table = pmf_table * ((rowsum - knum - jnum) * (colsum(2) - jnum))
            jnum = jnum + 1#
            pmf_table = pmf_table / (jnum * (col1mRowSum + knum + jnum))
            If pmf_table = 0# Then
               ccdf = 0#
               temp = 0#
               Exit Do
            End If
            If pmf_table * pmfh <= pmf_Obs Then Exit Do
            temp = temp - pmf_table
         Loop
         jnum = jnum - 1#
      End If
      If k = 50 Then pmf_table_jnum = pmf_hypergeometric(jnum, rowsum - knum, colsum(2), colsum12)
      ccdf = ccdf + temp
      If inum > jnum Then Exit Do
      Call AddValueToStack(ast, pmfh * (cdf + ccdf))
'Debug.Print knum, pmfh * (cdf + ccdf)
   Loop
   If pmfh > 0# Then
      prob = cdf_hypergeometric(knum, rowsum, colsum(3), cs)
   End If
   inum = inumstart
   jnum = jnumstart
   cdf = cdf_save
   ccdf = ccdf_save
   pmf_table_inum = pmf_table_inum_save
   pmf_table_jnum = pmf_table_jnum_save
   knum = ml(3)
   pmfh = pmfh_save
   k = 0
   Do While knum <= colsum(3)
      k = k + 1
      If k = 100 Then
         knum = knum + 1#
         pmfh = pmf_hypergeometric(knum, rowsum, colsum(3), cs)
         k = 0
      Else
         pmfh = pmfh * ((rowsum - knum) * (colsum(3) - knum))
         knum = knum + 1#
         pmfh = pmfh / (knum * (colsum12 - rowsum + knum))
         'pmfh = pmf_hypergeometric(knum, rowsum, colsum(3), cs)
      End If
      If pmfh <= pmf_Obs Then Exit Do
      mode = Int((rowsum - knum + 1#) * (colsum(2) + 1#) / (colsum12 + 2#))
      pmf_table = pmf_table_inum * ((rowsum - knum - inum + 1#) * (colsum12 - rowsum + knum)) / ((rowsum - knum + 1#) * (col1mRowSum + knum + inum))   'pmf_hypergeometric(inum, rowsum - knum, colsum(2), colsum12)
      temp = pmf_table_inum * inum / (rowsum - knum + 1#)
      If pmf_table = 0# Then
         inum = inum - 1#
         pmf_table = temp * (colsum12 - rowsum + knum) / (colsum(2) - inum) 'pmf_hypergeometric(inum, rowsum - knum, colsum(2), colsum12)
         temp = temp - pmf_table
      End If
      Do While inum > mode
         pmf_table = pmf_table * (inum * (col1mRowSum + knum + inum))
         inum = inum - 1#
         pmf_table = pmf_table / ((rowsum - knum - inum) * (colsum(2) - inum))
         temp = temp - pmf_table
      Loop
      If pmf_table * pmfh <= pmf_Obs Then
         Do
            temp = temp + pmf_table
            pmf_table = pmf_table * ((rowsum - knum - inum) * (colsum(2) - inum))
            inum = inum + 1#
            pmf_table = pmf_table / (inum * (col1mRowSum + knum + inum))
         Loop Until (pmf_table * pmfh > pmf_Obs) Or (inum > mode)
         pmf_table_inum = pmf_table
      Else 'If cdf > 0# Then
         Do
            pmf_table_inum = pmf_table
            pmf_table = pmf_table * (inum * (col1mRowSum + knum + inum))
            inum = inum - 1#
            pmf_table = pmf_table / ((rowsum - knum - inum) * (colsum(2) - inum))
            If pmf_table = 0# Then
               cdf = 0#
               temp = 0#
               Exit Do
            End If
            If pmf_table * pmfh <= pmf_Obs Then Exit Do
            temp = temp - pmf_table
         Loop
         inum = inum + 1#
      End If
      If k = 50 Then pmf_table_inum = pmf_hypergeometric(inum, rowsum - knum, colsum(2), colsum12)
      cdf = cdf + temp
      jnum = jnum - 1#
      pmf_table = pmf_table_jnum * ((jnum + 1#) * (colsum12 - rowsum + knum)) / ((rowsum - knum + 1#) * (colsum(2) - jnum))    'pmf_hypergeometric(jnum, rowsum - knum, colsum(2), colsum12)
      temp = pmf_table_jnum * (rowsum - knum - jnum) / (rowsum - knum + 1#)
      If pmf_table = 0# Then
         jnum = jnum + 1#
         pmf_table = temp * (colsum(1) - rowsum + knum - jnum) / (colsum12 - rowsum + knum)  'pmf_hypergeometric(jnum, rowsum - knum, colsum(2), colsum12)
         temp = temp - pmf_table
      End If
      Do While jnum < mode ' And pmf_table > 0#
         pmf_table = pmf_table * ((rowsum - knum - jnum) * (colsum(2) - jnum))
         jnum = jnum + 1#
         pmf_table = pmf_table / (jnum * (col1mRowSum + knum + jnum))
         temp = temp - pmf_table
      Loop
      If pmf_table * pmfh <= pmf_Obs Then
         Do
            temp = temp + pmf_table
            pmf_table = pmf_table * (jnum * (col1mRowSum + knum + jnum))
            jnum = jnum - 1#
            pmf_table = pmf_table / ((rowsum - knum - jnum) * (colsum(2) - jnum))
         Loop Until (pmf_table * pmfh > pmf_Obs) Or (jnum < mode)
         pmf_table_jnum = pmf_table
      Else 'If ccdf > 0# Then
         Do
            pmf_table_jnum = pmf_table
            pmf_table = pmf_table * ((rowsum - knum - jnum) * (colsum(2) - jnum))
            jnum = jnum + 1#
            pmf_table = pmf_table / (jnum * (col1mRowSum + knum + jnum))
            If pmf_table = 0# Then
               ccdf = 0#
               temp = 0#
               Exit Do
            End If
            If pmf_table * pmfh <= pmf_Obs Then Exit Do
            temp = temp - pmf_table
         Loop
         jnum = jnum - 1#
      End If
      If k = 50 Then pmf_table_jnum = pmf_hypergeometric(jnum, rowsum - knum, colsum(2), colsum12)
      ccdf = ccdf + temp
      If inum > jnum Then Exit Do
      Call AddValueToStack(ast, pmfh * (cdf + ccdf))
'Debug.Print knum, pmfh * (cdf + ccdf)
   Loop
   If pmfh > 0# Then
      prob = prob + comp_cdf_hypergeometric(knum - 1#, rowsum, colsum(3), cs)
   End If
   fet_23 = prob + StackTotal(ast)
'Call DumpAddStack(ast)
'Debug.Print fet_23
Else
   prob = 0#
   Call InitAddStack(ast)
   inum = inumstart
   jnum = jnumstart
   prob_d = fet_23(c - 1, colsum, rowsum - knum, pmf_Obs / pmfh, inum, jnum)
'Debug.Print knum, pmfh, prob_d, pmfh * prob_d
   Call AddValueToStack(ast, pmfh * prob_d)
   inumstart = inum
   jnumstart = jnum
   If inum > jnum Then
      fet_23 = 1#
      Exit Function
   End If
   Do
      pmfh = pmfh * ((colsum(c) - knum) * (rowsum - knum))
      knum = knum + 1#
      pmfh = pmfh / (knum * (cs - colsum(c) - rowsum + knum))
      If pmfh = 0# Then Exit Do
      prob_d = fet_23(c - 1, colsum, rowsum - knum, pmf_Obs / pmfh, inum, jnum)
      If inum > jnum Then
         prob = prob + comp_cdf_hypergeometric(knum - 1#, colsum(c), rowsum, cs)
         Exit Do
      End If
      Call AddValueToStack(ast, pmfh * prob_d)
   Loop
   pmfh = pmfh_save
   knum = ml(c)
   inum = inumstart
   jnum = jnumstart
   Do
      pmfh = pmfh * (knum * (cs - colsum(c) - rowsum + knum))
      knum = knum - 1#
      pmfh = pmfh / ((colsum(c) - knum) * (rowsum - knum))
      If pmfh = 0# Then Exit Do
      prob_d = fet_23(c - 1, colsum, rowsum - knum, pmf_Obs / pmfh, inum, jnum)
      If inum > jnum Then
         prob = prob + cdf_hypergeometric(knum, colsum(c), rowsum, cs)
         Exit Do
      End If
      Call AddValueToStack(ast, pmfh * prob_d)
   Loop
   fet_23 = prob + StackTotal(ast)
End If
End Function

Function fet_24(c As Long, ByRef colsum() As Double, ByVal rowsum As Double, ByVal pmf_Obs As Double, ByRef inumstart As Double, ByRef jnumstart As Double) As Double
Dim d As Double, cs As Double, colsum12 As Double, colsum34 As Double, prob As Double, temp As Double, inum As Double, jnum As Double, inum_save As Double, jnum_save As Double
Dim cdf As Double, ccdf As Double, pmf_table_inum As Double, pmf_table_jnum As Double, mode As Double, cdf_save As Double, ccdf_save As Double, col1mRowSum As Double, probf4 As Double
Dim dnum_old As Double, fnum_old As Double, pmfd_old As Double, dnum_save As Double, pmfd_save As Double
Dim i As Long, j As Long, count As Long
Dim all_d_zero As Boolean
Dim asto As TAddStack

ReDim ml(1 To c) As Double
If pmf_Obs > 1# Then 'All tables have pmf <= 1
   fet_24 = 1#
   Exit Function
End If

cs = 0#
For i = 1 To c
   cs = cs + colsum(i)
Next i
'First guess at mode vector
ml(1) = rowsum
For i = 2 To c
   ml(i) = Int(rowsum * colsum(i) / cs + 0.5)
   ml(1) = ml(1) - ml(i)
Next i

Do 'Update guess at mode vector
   all_d_zero = True
   For i = 1 To c - 1
      For j = i + 1 To c
         d = ml(i) - Int((colsum(i) + 1#) * (ml(i) + ml(j) + 1#) / (colsum(i) + colsum(j) + 2#))
         If d <> 0# Then
            ml(i) = ml(i) - d
            ml(j) = ml(j) + d
            all_d_zero = False
         End If
      Next j
   Next i
Loop Until all_d_zero

Dim pmff As Double, pmfd As Double, pmfd_down As Double, pmfd_up As Double
Dim cdff As Double, probf As Double, pmf_Obs_save As Double
Dim dnum As Double, dnum_up As Double, dnum_down As Double, fnum As Double, fnum_save As Double, pmff_save As Double, rowsummfnum As Double
Dim cdf_start As Double, ccdf_start As Double, inum_min As Double, jnum_max As Double
Dim continue_up As Boolean, continue_down As Boolean, exit_loop As Boolean
If c = 4 Then
   Call InitAddStack(asto)
   prob = 0#
   dnum = ml(4)
   fnum = ml(3) + dnum
   rowsummfnum = rowsum - fnum
   fnum_save = fnum
   colsum12 = colsum(1) + colsum(2)
   colsum34 = colsum(3) + colsum(4)
   pmff = pmf_hypergeometric(fnum, colsum34, rowsum, cs)
   pmff_save = pmff
   pmfd = pmf_hypergeometric(dnum, fnum, colsum(4), colsum34)
   pmf_Obs_save = pmf_Obs
   pmf_Obs = pmf_Obs_save / (pmfd * pmff)
   col1mRowSum = colsum(1) - rowsum
   inum_min = Max(0#, -(fnum + col1mRowSum))
   inum = Min(Max(inum_min, inumstart), ml(2))
   jnum_max = Min(rowsummfnum, colsum(2))
   jnum = Max(Min(jnum_max, jnumstart), ml(2))
   pmf_table_inum = pmf_hypergeometric(inum, rowsummfnum, colsum(2), colsum12)
   mode = Int((rowsummfnum + 1#) * (colsum(2) + 1#) / (colsum12 + 2#))
   Do While pmf_table_inum = 0#
      inum = Int((inum + Max(2, mode)) * 0.5)
      pmf_table_inum = pmf_hypergeometric(inum, rowsummfnum, colsum(2), colsum12)
   Loop
   Do While (pmf_table_inum > pmf_Obs) And (inum > inum_min)
      pmf_table_inum = pmf_table_inum * (inum * (col1mRowSum + fnum + inum))
      inum = inum - 1#
      pmf_table_inum = pmf_table_inum / ((rowsummfnum - inum) * (colsum(2) - inum))
   Loop
   Do While (pmf_table_inum <= pmf_Obs) And (inum <= mode)
      pmf_table_inum = pmf_table_inum * (rowsummfnum - inum) * (colsum(2) - inum)
      inum = inum + 1#
      pmf_table_inum = pmf_table_inum / (inum * (col1mRowSum + fnum + inum))
   Loop
   pmf_table_jnum = pmf_hypergeometric(jnum, rowsummfnum, colsum(2), colsum12)
   Do While pmf_table_jnum = 0#
      jnum = Int((jnum + mode) * 0.5)
      pmf_table_jnum = pmf_hypergeometric(jnum, rowsummfnum, colsum(2), colsum12)
   Loop
   Do While (pmf_table_jnum > pmf_Obs) And (jnum < jnum_max)
      pmf_table_jnum = pmf_table_jnum * ((rowsummfnum - jnum) * (colsum(2) - jnum))
      jnum = jnum + 1#
      pmf_table_jnum = pmf_table_jnum / (jnum * (col1mRowSum + fnum + jnum))
   Loop
   Do While (pmf_table_jnum < pmf_Obs) And (jnum >= mode)
      pmf_table_jnum = pmf_table_jnum * (jnum * (col1mRowSum + fnum + jnum))
      jnum = jnum - 1#
      pmf_table_jnum = pmf_table_jnum / ((rowsummfnum - jnum) * (colsum(2) - jnum))
   Loop
   inumstart = inum
   jnumstart = jnum
   If inum > jnum Then
      fet_24 = 1#
      Exit Function
   End If
   cdf = cdf_hypergeometric(inum - 1, rowsummfnum, colsum(2), colsum12)
   ccdf = comp_cdf_hypergeometric(jnum, rowsummfnum, colsum(2), colsum12)
   cdf_start = cdf
   ccdf_start = ccdf
   continue_up = True
   continue_down = True
   inum_save = inum
   jnum_save = jnum
   cdf_save = cdf
   ccdf_save = ccdf
   cdf_save = cdf
   ccdf_save = ccdf
   prob = 0#
   dnum_save = dnum
   pmfd_save = pmfd
   Do
      dnum_old = dnum
      fnum_old = fnum
      pmfd_old = pmfd
      probf4 = 0#
      count = 1
      probf = pmfd * (cdf + ccdf)
      pmf_Obs = pmf_Obs * pmfd ' was pmf_Obs = pmf_Obs / (pmfd * pmff)
      pmfd_down = pmfd
      pmfd_up = pmfd
      dnum_down = dnum
      dnum_up = dnum
      pmfd_down = pmfd_down * (dnum_down * (colsum(3) - fnum + dnum_down))
      dnum_down = dnum_down - 1
      pmfd_down = pmfd_down / ((fnum - dnum_down) * (colsum(4) - dnum_down))
      pmfd_up = pmfd_up * ((fnum - dnum_up) * (colsum(4) - dnum_up))
      dnum_up = dnum_up + 1#
      pmfd_up = pmfd_up / (dnum_up * (colsum(3) - fnum + dnum_up))
      Do
         pmfd = Max(pmfd_down, pmfd_up)
         If pmfd = 0# Then Exit Do
         Do While pmfd * pmf_table_inum <= pmf_Obs
            cdf = cdf + pmf_table_inum
            pmf_table_inum = pmf_table_inum * (rowsummfnum - inum) * (colsum(2) - inum)
            inum = inum + 1#
            pmf_table_inum = pmf_table_inum / (inum * (col1mRowSum + fnum + inum))
            If (inum > mode) Then Exit Do
         Loop
         Do While pmfd * pmf_table_jnum <= pmf_Obs
            ccdf = ccdf + pmf_table_jnum
            pmf_table_jnum = pmf_table_jnum * (jnum * (col1mRowSum + fnum + jnum))
            jnum = jnum - 1#
            pmf_table_jnum = pmf_table_jnum / ((rowsummfnum - jnum) * (colsum(2) - jnum))
            If (jnum < mode) Then Exit Do
         Loop
         If inum > jnum Then Exit Do
         probf4 = probf4 + (pmfd * (cdf + ccdf) - probf)
         count = count + 1
         If pmfd_down > pmfd_up Then
            pmfd_down = pmfd_down * (dnum_down * (colsum(3) - fnum + dnum_down))
            dnum_down = dnum_down - 1#
            pmfd_down = pmfd_down / ((fnum - dnum_down) * (colsum(4) - dnum_down))
         Else
            pmfd_up = pmfd_up * ((fnum - dnum_up) * (colsum(4) - dnum_up))
            dnum_up = dnum_up + 1#
            pmfd_up = pmfd_up / (dnum_up * (colsum(3) - fnum + dnum_up))
         End If
'Debug.Print inum, jnum, pmf_table_inum, pmf_table_jnum, cdf, ccdf, probf
      Loop
      probf = count * probf + probf4
      If inum > jnum Then
         pmfd_down = cdf_hypergeometric(dnum_down, fnum, colsum(4), colsum34)
         pmfd_up = comp_cdf_hypergeometric(dnum_up - 1#, fnum, colsum(4), colsum34)
         probf = probf + pmfd_down + pmfd_up
      End If

      Call AddValueToStack(asto, probf * pmff)
      If continue_up Then
         pmff = pmff * ((rowsummfnum) * (colsum34 - fnum))
         fnum = fnum + 1#
         rowsummfnum = rowsum - fnum
         pmff = pmff / ((colsum12 - rowsummfnum) * (fnum))
         continue_up = pmff > 0#
         If Not continue_up Then
            pmff = pmff_save
            fnum = fnum_save
            rowsummfnum = rowsum - fnum
            dnum_old = dnum_save
            fnum_old = fnum_save
            pmfd_old = pmfd_save
            inum_save = inumstart
            jnum_save = jnumstart
            cdf_save = cdf_start
            ccdf_save = ccdf_start
         Else
            inum = inum_save
            jnum = jnum_save
            cdf = cdf_save
            ccdf = ccdf_save
            inum = inum - 1#
            temp = PBB(col1mRowSum + fnum + inum, rowsummfnum - inum, colsum(2) - inum, inum + 1#)
            hTerm = hTerm * ((rowsummfnum - inum) * (colsum(2) - inum) * (colsum12 + 1#))
            inum = inum + 1#
            hTerm = hTerm / ((rowsummfnum + 1#) * (colsum(2) + 1#) * (col1mRowSum + fnum + inum))
            pmf_table_inum = hTerm 'pmf_hypergeometric(inum, rowsummfnum, colsum(2), colsum12)
            cdf = cdf + temp
            temp = PBB(jnum, colsum(2) - jnum, rowsummfnum - jnum + 1#, col1mRowSum + fnum + jnum)
            hTerm = hTerm * (jnum * (col1mRowSum + fnum + jnum) * (colsum12 + 1#))
            jnum = jnum - 1#
            hTerm = hTerm / ((rowsummfnum + 1#) * (colsum(2) - jnum) * (colsum(1) + 1#))
            pmf_table_jnum = hTerm 'pmf_hypergeometric(jnum, rowsummfnum, colsum(2), colsum12)
            ccdf = ccdf + temp
         End If
      End If
      Do
         If Not continue_up Then
            pmff = pmff * ((colsum12 - rowsummfnum) * (fnum))
            fnum = fnum - 1#
            rowsummfnum = rowsum - fnum
            pmff = pmff / ((rowsummfnum) * (colsum34 - fnum))
            continue_down = pmff > 0#
            If Not continue_down Then Exit Do
            inum = inum_save
            jnum = jnum_save
            cdf = cdf_save
            ccdf = ccdf_save
            temp = PBB(inum, colsum(2) - inum, rowsummfnum - inum, col1mRowSum + fnum + inum + 1#)
            hTerm = hTerm * ((rowsummfnum - inum) * (colsum(2) - inum) * (colsum12 + 1#))
            inum = inum + 1#
            hTerm = hTerm / (inum * (colsum(1) + 1#) * (colsum12 - rowsummfnum + 1#))
            pmf_table_inum = hTerm 'pmf_hypergeometric(inum, rowsummfnum, colsum(2), colsum12)
            cdf = cdf + temp
            temp = PBB(col1mRowSum + fnum + jnum + 1#, rowsummfnum - jnum - 1#, colsum(2) - jnum, jnum + 1#)
            hTerm = hTerm * ((jnum + 1#) * (col1mRowSum + fnum + jnum + 1#) * (colsum12 + 1#)) / ((colsum(2) + 1#) * (rowsummfnum - jnum) * (colsum12 - rowsummfnum + 1#))
            pmf_table_jnum = hTerm 'pmf_hypergeometric(jnum, rowsummfnum, colsum(2), colsum12)
            ccdf = ccdf + temp
         End If
'Initialise for next go round loop
         mode = Int((rowsummfnum + 1#) * (colsum(2) + 1#) / (colsum12 + 2#))
         dnum = Int((fnum + 1#) * (colsum(4) + 1#) / (colsum34 + 2#))
         If continue_up Then
            If dnum = dnum_old Then
               pmfd = pmfd_old * ((colsum34 - fnum_old - colsum(4) + dnum) * (fnum)) / ((fnum - dnum) * (colsum34 - fnum_old))
            ElseIf dnum = dnum_old + 1# Then
               pmfd = pmfd_old * (fnum * (colsum(4) - dnum_old)) / (dnum * (colsum34 - fnum_old))
            Else
               Debug.Print dnum_old, fnum_old, dnum, fnum
               pmfd = 1# / 0#
               pmfd = pmf_hypergeometric(dnum, fnum, colsum(4), colsum34)
            End If
         Else
            If dnum = dnum_old Then
               pmfd = pmfd_old * ((fnum_old - dnum) * (colsum34 - fnum)) / ((colsum34 - fnum - colsum(4) + dnum) * (fnum_old))
            ElseIf dnum = dnum_old - 1# Then
               pmfd = pmfd_old * (dnum_old * (colsum34 - fnum)) / (fnum_old * (colsum(4) - dnum))
            Else
               Debug.Print dnum_old, fnum_old, dnum, fnum
               pmfd = 1# / 0#
               pmfd = pmf_hypergeometric(dnum, fnum, colsum(4), colsum34)
            End If
         End If
         'pmfd = pmf_hypergeometric(dnum, fnum, colsum(4), colsum34)
         pmf_Obs = pmf_Obs_save / (pmff * pmfd)
         inum_min = Max(0#, -(fnum + col1mRowSum))
         jnum_max = Min(rowsummfnum, colsum(2))
         If jnum_max >= 0# And pmf_Obs < 1# Then
            Do While (pmf_table_inum > pmf_Obs) And (inum > inum_min) Or (inum > mode)
               pmf_table_inum = pmf_table_inum * (inum * (col1mRowSum + fnum + inum))
               inum = inum - 1#
               pmf_table_inum = pmf_table_inum / ((rowsummfnum - inum) * (colsum(2) - inum))
               cdf = cdf - pmf_table_inum
            Loop
            Do While (pmf_table_inum < pmf_Obs) And (inum <= mode)
               cdf = cdf + pmf_table_inum
               pmf_table_inum = pmf_table_inum * (rowsummfnum - inum) * (colsum(2) - inum)
               inum = inum + 1#
               pmf_table_inum = pmf_table_inum / (inum * (col1mRowSum + fnum + inum))
            Loop
            If inum = inum_min Then cdf = 0#
            Do While (pmf_table_jnum > pmf_Obs) And (jnum < jnum_max) Or (jnum < mode)
               pmf_table_jnum = pmf_table_jnum * ((rowsummfnum - jnum) * (colsum(2) - jnum))
               jnum = jnum + 1#
               pmf_table_jnum = pmf_table_jnum / (jnum * (col1mRowSum + fnum + jnum))
               ccdf = ccdf - pmf_table_jnum
            Loop
            Do While (pmf_table_jnum < pmf_Obs) And (jnum >= mode)
               ccdf = ccdf + pmf_table_jnum
               pmf_table_jnum = pmf_table_jnum * (jnum * (col1mRowSum + fnum + jnum))
               jnum = jnum - 1#
               pmf_table_jnum = pmf_table_jnum / ((rowsummfnum - jnum) * (colsum(2) - jnum))
            Loop
            If jnum = jnum_max Then ccdf = 0#
         Else
            inum = jnum + 1#
         End If
         exit_loop = True
         If inum > jnum Then
            If continue_up Then
               prob = prob + comp_cdf_hypergeometric(fnum - 1#, colsum34, rowsum, cs)
               continue_up = False
               pmff = pmff_save
               fnum = fnum_save
               rowsummfnum = rowsum - fnum
               dnum_old = dnum_save
               fnum_old = fnum_save
               pmfd_old = pmfd_save
               inum_save = inumstart
               jnum_save = jnumstart
               cdf_save = cdf_start
               ccdf_save = ccdf_start
               exit_loop = False
            Else
               prob = prob + cdf_hypergeometric(fnum, colsum34, rowsum, cs)
               continue_down = False
            End If
         End If
      Loop Until exit_loop
      If Not continue_down Then Exit Do
      cdf_save = cdf
      ccdf_save = ccdf
      inum_save = inum
      jnum_save = jnum
   Loop Until Not continue_down
   fet_24 = prob + StackTotal(asto)
ElseIf c > 4 Then
   Dim knum As Double
   knum = ml(c)
   pmff = pmf_hypergeometric(knum, colsum(c), rowsum, cs)
   pmff_save = pmff
   inum = inumstart
   jnum = jnumstart
   Call InitAddStack(asto)
   probf4 = fet_24(c - 1, colsum, rowsum - knum, pmf_Obs / pmff, inum, jnum)
   inumstart = inum
   jnumstart = jnum
   If inum > jnum Then
      fet_24 = 1#
      Exit Function
   End If
   Call AddValueToStack(asto, pmff * probf4)

   Do
      pmff = pmff * ((colsum(c) - knum) * (rowsum - knum))
      knum = knum + 1#
      pmff = pmff / (knum * (cs - colsum(c) - rowsum + knum))
      If pmff = 0# Then Exit Do
      probf4 = fet_24(c - 1, colsum, rowsum - knum, pmf_Obs / pmff, inum, jnum)
      If inum > jnum Then
         prob = prob + comp_cdf_hypergeometric(knum - 1#, colsum(c), rowsum, cs)
         Exit Do
      End If
      Call AddValueToStack(asto, pmff * probf4)
   Loop
   pmff = pmff_save
   knum = ml(c)
   inum = inumstart
   jnum = jnumstart
   Do
      pmff = pmff * (knum * (cs - colsum(c) - rowsum + knum))
      knum = knum - 1#
      pmff = pmff / ((colsum(c) - knum) * (rowsum - knum))
      If pmff = 0# Then Exit Do
      probf4 = fet_24(c - 1, colsum, rowsum - knum, pmf_Obs / pmff, inum, jnum)
      If inum > jnum Then
         prob = prob + cdf_hypergeometric(knum, colsum(c), rowsum, cs)
         Exit Do
      End If
      Call AddValueToStack(asto, pmff * probf4)
   Loop
   fet_24 = prob + StackTotal(asto)
Else
   fet_24 = "Values other than 4 not allowed"
End If
End Function

Function fet_25(c As Long, ByRef colsum() As Double, ByVal rowsum As Double, ByVal pmf_Obs As Double, ByRef inumstart As Double, ByRef jnumstart As Double) As Double
Dim d As Double, cs As Double, prob As Double, temp As Double, inum As Double, jnum As Double, inum_save As Double, jnum_save As Double, inum_save_save As Double, jnum_save_save As Double
Dim cdf As Double, ccdf As Double, pmf_table_inum As Double, pmf_table_jnum As Double, mode As Double, cdf_save As Double, ccdf_save As Double, col1mRowSum As Double, prob_d As Double
Dim i As Long, j As Long, c_count As Long
Dim all_d_zero As Boolean

ReDim ml(1 To c) As Double

If pmf_Obs > 1# Then 'All tables have pmf <= 1
   fet_25 = 1#
   Exit Function
End If

cs = 0#
For i = 1 To c
   cs = cs + colsum(i)
Next i
'First guess at mode vector
ml(1) = rowsum
For i = 2 To c
   ml(i) = Int(rowsum * colsum(i) / cs + 0.5)
   ml(1) = ml(1) - ml(i)
Next i

Do 'Update guess at mode vector
   all_d_zero = True
   For i = 1 To c - 1
      For j = i + 1 To c
         d = ml(i) - Int((colsum(i) + 1#) * (ml(i) + ml(j) + 1#) / (colsum(i) + colsum(j) + 2#))
         If d <> 0# Then
            ml(i) = ml(i) - d
            ml(j) = ml(j) + d
            all_d_zero = False
         End If
      Next j
   Next i
Loop Until all_d_zero

Dim pmff As Double, pmfd4 As Double, pmfd4_down As Double, pmfd4_up As Double, pmfd5 As Double, pmfd5_save As Double
Dim cdff As Double, probf As Double, pmf_Obs_save As Double, probf5 As Double, cdf_save_save As Double, ccdf_save_save As Double
Dim d4num As Double, d4num_up As Double, d4num_down As Double, d5num As Double, d5num_save As Double, fnum As Double, fnum_save As Double, pmff_save As Double, rowsummfnum As Double
Dim cdf_start As Double, ccdf_start As Double, inum_min As Double, jnum_max As Double, pmf_table_inum_save As Double, pmf_table_jnum_save As Double
Dim inum_save5 As Double, jnum_save5 As Double, pmf_table_inum_save5 As Double, pmf_table_jnum_save5 As Double, cdf_save5 As Double, ccdf_save5 As Double
Dim colsum12 As Double, colsum34 As Double, colsum345 As Double

Dim continue_up As Boolean, continue_down As Boolean, exit_loop As Boolean
Dim c5_up As Boolean, c5_down As Boolean, el5 As Boolean
If c = 5 Then
   d5num = ml(5)
   d4num = ml(4)
   fnum = ml(3) + d4num + d5num
   rowsummfnum = rowsum - fnum
   fnum_save = fnum
   colsum12 = colsum(1) + colsum(2)
   colsum34 = colsum(3) + colsum(4)
   colsum345 = colsum34 + colsum(5)
   pmff = pmf_hypergeometric(fnum, colsum345, rowsum, cs)
   pmff_save = pmff
   pmfd4 = pmf_hypergeometric(d4num, fnum - d5num, colsum(4), colsum34)
   pmfd5 = pmf_hypergeometric(d5num, fnum, colsum(5), colsum345)
   pmf_Obs_save = pmf_Obs
   pmf_Obs = pmf_Obs_save / (pmfd4 * pmfd5 * pmff)
   col1mRowSum = colsum(1) - rowsum
   inum_min = Max(0#, -(fnum + col1mRowSum))
   inum = Min(Max(inum_min, inumstart), ml(2))
   jnum_max = Min(rowsummfnum, colsum(2))
   jnum = Max(Min(jnum_max, jnumstart), ml(2))
   pmf_table_inum = pmf_hypergeometric(inum, rowsummfnum, colsum(2), colsum12)
   mode = Int((rowsummfnum + 1#) * (colsum(2) + 1#) / (colsum12 + 2#))
   Do While pmf_table_inum = 0#
      inum = Int((inum + Max(2, mode)) * 0.5)
      pmf_table_inum = pmf_hypergeometric(inum, rowsummfnum, colsum(2), colsum12)
   Loop
   Do While (pmf_table_inum > pmf_Obs) And (inum > inum_min)
      pmf_table_inum = pmf_table_inum * (inum * (col1mRowSum + fnum + inum))
      inum = inum - 1#
      pmf_table_inum = pmf_table_inum / ((rowsummfnum - inum) * (colsum(2) - inum))
   Loop
   Do While (pmf_table_inum <= pmf_Obs) And (inum <= mode)
      pmf_table_inum = pmf_table_inum * (rowsummfnum - inum) * (colsum(2) - inum)
      inum = inum + 1#
      pmf_table_inum = pmf_table_inum / (inum * (col1mRowSum + fnum + inum))
   Loop
   pmf_table_jnum = pmf_hypergeometric(jnum, rowsummfnum, colsum(2), colsum12)
   Do While pmf_table_jnum = 0#
      jnum = Int((jnum + mode) * 0.5)
      pmf_table_jnum = pmf_hypergeometric(jnum, rowsummfnum, colsum(2), colsum12)
   Loop
   Do While (pmf_table_jnum > pmf_Obs) And (jnum < jnum_max)
      pmf_table_jnum = pmf_table_jnum * ((rowsummfnum - jnum) * (colsum(2) - jnum))
      jnum = jnum + 1#
      pmf_table_jnum = pmf_table_jnum / (jnum * (col1mRowSum + fnum + jnum))
   Loop
   Do While (pmf_table_jnum < pmf_Obs) And (jnum >= mode)
      pmf_table_jnum = pmf_table_jnum * (jnum * (col1mRowSum + fnum + jnum))
      jnum = jnum - 1#
      pmf_table_jnum = pmf_table_jnum / ((rowsummfnum - jnum) * (colsum(2) - jnum))
   Loop
   inumstart = inum
   jnumstart = jnum
   If inum > jnum Then
      fet_25 = 1#
      Exit Function
   End If
   cdf = cdf_hypergeometric(inum - 1, rowsummfnum, colsum(2), colsum12)
   ccdf = comp_cdf_hypergeometric(jnum, rowsummfnum, colsum(2), colsum12)
   cdf_start = cdf
   ccdf_start = ccdf
   continue_up = True
   continue_down = True
   inum_save = inum
   jnum_save = jnum
   inum_save_save = inum
   jnum_save_save = jnum
   pmf_table_inum_save = pmf_table_inum
   pmf_table_jnum_save = pmf_table_jnum
   cdf_save = cdf
   ccdf_save = ccdf
   cdf_save_save = cdf
   ccdf_save_save = ccdf
   prob = 0#
   probf5 = 0#
   Do
      d5num_save = d5num
      pmfd5_save = pmfd5
      inum_save5 = inum
      jnum_save5 = jnum
      pmf_table_inum_save5 = pmf_table_inum
      pmf_table_jnum_save5 = pmf_table_jnum
      cdf_save5 = cdf
      ccdf_save5 = ccdf
      c5_up = True
      c5_down = True
      Do
         probf = pmfd4 * (cdf + ccdf)
         pmf_Obs = pmf_Obs * pmfd4 ' was pmf_Obs = pmf_Obs / (pmfd4 * pmfd5 * pmff)
         pmfd4_down = pmfd4
         pmfd4_up = pmfd4
         d4num_down = d4num
         d4num_up = d4num
         pmfd4_down = pmfd4_down * d4num_down * (colsum(3) - fnum + d5num + d4num_down)
         d4num_down = d4num_down - 1
         pmfd4_down = pmfd4_down / ((fnum - d5num - d4num_down) * (colsum(4) - d4num_down))
         pmfd4_up = pmfd4_up * ((fnum - d5num - d4num_up) * (colsum(4) - d4num_up))
         d4num_up = d4num_up + 1#
         pmfd4_up = pmfd4_up / (d4num_up * (colsum(3) - fnum + d5num + d4num_up))
         Do
            pmfd4 = Max(pmfd4_down, pmfd4_up)
            If pmfd4 = 0# Then Exit Do
            If pmfd4 * pmf_table_inum <= pmf_Obs Then
               Do
                  cdf = cdf + pmf_table_inum
                  pmf_table_inum = pmf_table_inum * (rowsummfnum - inum) * (colsum(2) - inum)
                  inum = inum + 1#
                  pmf_table_inum = pmf_table_inum / (inum * (col1mRowSum + fnum + inum))
               Loop Until (pmfd4 * pmf_table_inum > pmf_Obs) Or (inum > mode)
            End If
            If pmfd4 * pmf_table_jnum < pmf_Obs Then
               Do
                  ccdf = ccdf + pmf_table_jnum
                  pmf_table_jnum = pmf_table_jnum * (jnum * (col1mRowSum + fnum + jnum))
                  jnum = jnum - 1#
                  pmf_table_jnum = pmf_table_jnum / ((rowsummfnum - jnum) * (colsum(2) - jnum))
               Loop Until (pmfd4 * pmf_table_jnum > pmf_Obs) Or (jnum < mode)
            End If
            If inum > jnum Then
               pmfd4_down = cdf_hypergeometric(d4num_down, fnum - d5num, colsum(4), colsum34)
               pmfd4_up = comp_cdf_hypergeometric(d4num_up - 1#, fnum - d5num, colsum(4), colsum34)
               probf = probf + pmfd4_down + pmfd4_up
               Exit Do
            End If
            probf = probf + pmfd4 * (cdf + ccdf)
            If pmfd4_down > pmfd4_up Then
               pmfd4_down = pmfd4_down * d4num_down * (colsum(3) - fnum + d5num + d4num_down)
               d4num_down = d4num_down - 1#
               pmfd4_down = pmfd4_down / ((fnum - d5num - d4num_down) * (colsum(4) - d4num_down))
            Else
               pmfd4_up = pmfd4_up * ((fnum - d5num - d4num_up) * (colsum(4) - d4num_up))
               d4num_up = d4num_up + 1#
               pmfd4_up = pmfd4_up / (d4num_up * (colsum(3) - fnum + d5num + d4num_up))
            End If
'Debug.Print inum, jnum, pmf_table_inum, pmf_table_jnum, cdf, ccdf, probf
         Loop
         probf5 = probf5 + probf * pmfd5
         If c5_up Then
            pmfd5 = pmfd5 * ((fnum - d5num) * (colsum(5) - d5num))
            d5num = d5num + 1#
            pmfd5 = pmfd5 / (d5num * (colsum34 - fnum + d5num))
            c5_up = pmfd5 > 0#
            If Not c5_up Then
               pmfd5 = pmfd5_save
               d5num = d5num_save
               inum_save_save = inum_save5
               jnum_save_save = jnum_save5
               pmf_table_inum_save = pmf_table_inum_save5
               pmf_table_jnum_save = pmf_table_jnum_save5
               cdf_save_save = cdf_save5
               ccdf_save_save = ccdf_save5
            End If
         End If
         Do
            If Not c5_up Then
               pmfd5 = pmfd5 * d5num * (colsum34 - fnum + d5num)
               d5num = d5num - 1#
               pmfd5 = pmfd5 / ((fnum - d5num) * (colsum(5) - d5num))
               c5_down = pmfd5 > 0#
               If Not c5_down Then Exit Do
            End If
            d4num = Int((fnum - d5num + 1#) * (colsum(4) + 1#) / (colsum34 + 2#))
            pmfd4 = pmf_hypergeometric(d4num, fnum - d5num, colsum(4), colsum34)
            pmf_Obs = pmf_Obs_save / (pmfd4 * pmfd5 * pmff)
            inum = inum_save_save
            jnum = jnum_save_save
            pmf_table_inum = pmf_table_inum_save
            pmf_table_jnum = pmf_table_jnum_save
            cdf = cdf_save_save
            ccdf = ccdf_save_save
            'Do While (pmf_table_inum > pmf_Obs) And (inum > inum_min)
            '   pmf_table_inum = pmf_table_inum * (inum * (col1mRowSum + fnum + inum))
            '   inum = inum - 1#
            '   pmf_table_inum = pmf_table_inum / ((rowsummfnum - inum) * (colsum(2) - inum))
            '   cdf = cdf - pmf_table_inum
            'Loop
            If pmf_table_inum * (inum * (col1mRowSum + fnum + inum)) / ((rowsummfnum - inum + 1#) * (colsum(2) - inum + 1#)) > pmf_Obs Then
               fet_25 = "Problem with cdf"
            End If
            Do While (pmf_table_inum <= pmf_Obs) And (inum <= mode)
               cdf = cdf + pmf_table_inum
               pmf_table_inum = pmf_table_inum * (rowsummfnum - inum) * (colsum(2) - inum)
               inum = inum + 1#
               pmf_table_inum = pmf_table_inum / (inum * (col1mRowSum + fnum + inum))
            Loop
            'Do While (pmf_table_jnum > pmf_Obs) And (jnum < jnum_max)
            '   pmf_table_jnum = pmf_table_jnum * ((rowsummfnum - jnum) * (colsum(2) - jnum))
            '   jnum = jnum + 1#
            '   pmf_table_jnum = pmf_table_jnum / (jnum * (col1mRowSum + fnum + jnum))
            '   ccdf = ccdf - pmf_table_jnum
            'Loop
            If pmf_table_jnum * ((rowsummfnum - jnum) * (colsum(2) - jnum)) / ((jnum + 1#) * (col1mRowSum + fnum + jnum + 1#)) > pmf_Obs Then
               fet_25 = "Problem with ccdf"
            End If
            Do While (pmf_table_jnum < pmf_Obs) And (jnum >= mode)
               ccdf = ccdf + pmf_table_jnum
               pmf_table_jnum = pmf_table_jnum * (jnum * (col1mRowSum + fnum + jnum))
               jnum = jnum - 1#
               pmf_table_jnum = pmf_table_jnum / ((rowsummfnum - jnum) * (colsum(2) - jnum))
            Loop
            el5 = True
            If inum > jnum Then
               If c5_up Then
                  probf5 = probf5 + comp_cdf_hypergeometric(d5num - 1#, fnum, colsum(5), colsum345)
                  pmfd5 = pmfd5_save
                  d5num = d5num_save
                  inum_save_save = inum_save5
                  jnum_save_save = jnum_save5
                  pmf_table_inum_save = pmf_table_inum_save5
                  pmf_table_jnum_save = pmf_table_jnum_save5
                  cdf_save_save = cdf_save5
                  ccdf_save_save = ccdf_save5
                  el5 = False
                  c5_up = False
               Else
                  probf5 = probf5 + cdf_hypergeometric(d5num, fnum, colsum(5), colsum345)
                  c5_down = False
                  Exit Do
               End If
            End If
         Loop Until el5
         If Not c5_down Then Exit Do
         inum_save_save = inum
         jnum_save_save = jnum
         pmf_table_inum_save = pmf_table_inum
         pmf_table_jnum_save = pmf_table_jnum
         cdf_save_save = cdf
         ccdf_save_save = ccdf
      Loop
      prob = prob + probf5 * pmff
'Debug.Print prob
      probf5 = 0#
      If continue_up Then
         pmff = pmff * ((rowsummfnum) * (colsum345 - fnum))
         fnum = fnum + 1#
         rowsummfnum = rowsum - fnum
         pmff = pmff / ((colsum12 - rowsummfnum) * (fnum))
         continue_up = pmff > 0#
         If Not continue_up Then
            pmff = pmff_save
            fnum = fnum_save
            rowsummfnum = rowsum - fnum
            inum_save = inumstart
            jnum_save = jnumstart
            cdf_save = cdf_start
            ccdf_save = ccdf_start
         Else
            inum = inum_save
            jnum = jnum_save
            cdf = cdf_save
            ccdf = ccdf_save
            inum = inum - 1#
            temp = PBB(col1mRowSum + fnum + inum, rowsummfnum - inum, colsum(2) - inum, inum + 1#)
            hTerm = hTerm * ((rowsummfnum - inum) * (colsum(2) - inum) * (colsum12 + 1#))
            inum = inum + 1#
            hTerm = hTerm / ((rowsummfnum + 1#) * (colsum(2) + 1#) * (col1mRowSum + fnum + inum))
            pmf_table_inum = hTerm 'pmf_hypergeometric(inum, rowsummfnum, colsum(2), colsum12)
            cdf = cdf + temp
            temp = PBB(jnum, colsum(2) - jnum, rowsummfnum - jnum + 1#, col1mRowSum + fnum + jnum)
            hTerm = hTerm * (jnum * (col1mRowSum + fnum + jnum) * (colsum12 + 1#))
            jnum = jnum - 1#
            hTerm = hTerm / ((rowsummfnum + 1#) * (colsum(2) - jnum) * (colsum(1) + 1#))
            pmf_table_jnum = hTerm 'pmf_hypergeometric(jnum, rowsummfnum, colsum(2), colsum12)
            ccdf = ccdf + temp
         End If
      End If
      Do
         If Not continue_up Then
            pmff = pmff * ((colsum12 - rowsummfnum) * (fnum))
            fnum = fnum - 1#
            rowsummfnum = rowsum - fnum
            pmff = pmff / ((rowsummfnum) * (colsum345 - fnum))
            continue_down = pmff > 0#
            If Not continue_down Then Exit Do
            inum = inum_save
            jnum = jnum_save
            cdf = cdf_save
            ccdf = ccdf_save
            temp = PBB(inum, colsum(2) - inum, rowsummfnum - inum, col1mRowSum + fnum + inum + 1#)
            hTerm = hTerm * (rowsummfnum - inum) * (colsum(2) - inum) * (colsum12 + 1#)
            inum = inum + 1#
            hTerm = hTerm / (inum * (colsum(1) + 1#) * (colsum12 - rowsummfnum + 1#))
            pmf_table_inum = hTerm 'pmf_hypergeometric(inum, rowsummfnum, colsum(2), colsum12)
            cdf = cdf + temp
            temp = PBB(col1mRowSum + fnum + jnum + 1#, rowsummfnum - jnum - 1#, colsum(2) - jnum, jnum + 1#)
            hTerm = hTerm * ((jnum + 1#) * (col1mRowSum + fnum + jnum + 1#) * (colsum12 + 1#)) / ((colsum(2) + 1#) * (rowsummfnum - jnum) * (colsum12 - rowsummfnum + 1#))
            pmf_table_jnum = hTerm 'pmf_hypergeometric(jnum, rowsummfnum, colsum(2), colsum12)
            ccdf = ccdf + temp
         End If
         mode = Int((rowsummfnum + 1#) * (colsum(2) + 1#) / (colsum12 + 2#))
         ml(4) = Int(fnum * colsum(4) / (colsum345) + 0.5)
         ml(5) = Int(fnum * colsum(5) / (colsum345) + 0.5)
         ml(3) = fnum - ml(4) - ml(5)
         Do 'Update guess at mode vector
            all_d_zero = True
            For i = 3 To c - 1
               For j = i + 1 To c
                  d = ml(i) - Int((colsum(i) + 1#) * (ml(i) + ml(j) + 1#) / (colsum(i) + colsum(j) + 2#))
                  If d <> 0# Then
                     ml(i) = ml(i) - d
                     ml(j) = ml(j) + d
                     all_d_zero = False
                  End If
               Next j
            Next i
         Loop Until all_d_zero
'Debug.Print ml(3), ml(4), ml(5)
         d5num = ml(5)
         d4num = ml(4)
         'If ml(4) <> Int((fnum - d5num + 1#) * (colsum(4) + 1#) / (colsum34 + 2#)) Then
         '   fet_25 = "Problem with mode"
         'End If
         pmfd4 = pmf_hypergeometric(d4num, fnum - d5num, colsum(4), colsum34)
         pmfd5 = pmf_hypergeometric(d5num, fnum, colsum(5), colsum345)
         pmf_Obs = pmf_Obs_save / (pmfd4 * pmfd5 * pmff)
         inum_min = Max(0#, -(fnum + col1mRowSum))
         jnum_max = Min(rowsummfnum, colsum(2))
         If jnum_max > 0# And pmf_Obs < 1# Then
            Do While (pmf_table_inum > pmf_Obs) And (inum > inum_min) Or (inum > mode)
               pmf_table_inum = pmf_table_inum * (inum * (col1mRowSum + fnum + inum))
               inum = inum - 1#
               pmf_table_inum = pmf_table_inum / ((rowsummfnum - inum) * (colsum(2) - inum))
               cdf = cdf - pmf_table_inum
            Loop
            Do While (pmf_table_inum < pmf_Obs) And (inum <= mode)
               cdf = cdf + pmf_table_inum
               pmf_table_inum = pmf_table_inum * (rowsummfnum - inum) * (colsum(2) - inum)
               inum = inum + 1#
               pmf_table_inum = pmf_table_inum / (inum * (col1mRowSum + fnum + inum))
            Loop
            If inum = inum_min Then cdf = 0#
            Do While (pmf_table_jnum > pmf_Obs) And (jnum < jnum_max) Or (jnum < mode)
               pmf_table_jnum = pmf_table_jnum * ((rowsummfnum - jnum) * (colsum(2) - jnum))
               jnum = jnum + 1#
               pmf_table_jnum = pmf_table_jnum / (jnum * (col1mRowSum + fnum + jnum))
               ccdf = ccdf - pmf_table_jnum
            Loop
            Do While (pmf_table_jnum < pmf_Obs) And (jnum >= mode)
               ccdf = ccdf + pmf_table_jnum
               pmf_table_jnum = pmf_table_jnum * (jnum * (col1mRowSum + fnum + jnum))
               jnum = jnum - 1#
               pmf_table_jnum = pmf_table_jnum / ((rowsummfnum - jnum) * (colsum(2) - jnum))
            Loop
            If jnum = jnum_max Then ccdf = 0#
         Else
            inum = jnum + 1#
         End If
         exit_loop = True
         If inum > jnum Then
            If continue_up Then
               prob = prob + comp_cdf_hypergeometric(fnum - 1#, colsum345, rowsum, cs)
'Debug.Print prob
               continue_up = False
               pmff = pmff_save
               fnum = fnum_save
               rowsummfnum = rowsum - fnum
               inum_save = inumstart
               jnum_save = jnumstart
               cdf_save = cdf_start
               ccdf_save = ccdf_start
               exit_loop = False
            Else
               prob = prob + cdf_hypergeometric(fnum, colsum345, rowsum, cs)
'Debug.Print prob
               continue_down = False
            End If
         End If
      Loop Until exit_loop
      If Not continue_down Then Exit Do
      cdf_save = cdf
      ccdf_save = ccdf
      inum_save = inum
      jnum_save = jnum
      cdf_save_save = cdf
      ccdf_save_save = ccdf
      inum_save_save = inum
      jnum_save_save = jnum
      pmf_table_inum_save = pmf_table_inum
      pmf_table_jnum_save = pmf_table_jnum
   Loop Until Not continue_down
   fet_25 = prob
   Exit Function
ElseIf c >= 6 Then
   Dim knum As Double
   knum = ml(c)
   pmff = pmf_hypergeometric(knum, colsum(c), rowsum, cs)
   pmff_save = pmff
   inum = inumstart
   jnum = jnumstart
   prob_d = fet_25(c - 1, colsum, rowsum - knum, pmf_Obs / pmff, inum, jnum)
   inumstart = inum
   jnumstart = jnum
   If inum > jnum Then
      fet_25 = 1#
      Exit Function
   End If
   prob = pmff * prob_d

   Do
      pmff = pmff * ((colsum(c) - knum) * (rowsum - knum))
      knum = knum + 1#
      pmff = pmff / (knum * (cs - colsum(c) - rowsum + knum))
      If pmff = 0# Then Exit Do
      prob_d = fet_25(c - 1, colsum, rowsum - knum, pmf_Obs / pmff, inum, jnum)
      If inum > jnum Then
         prob = prob + comp_cdf_hypergeometric(knum - 1#, colsum(c), rowsum, cs)
         Exit Do
      End If
      prob = prob + pmff * prob_d
   Loop
   pmff = pmff_save
   knum = ml(c)
   inum = inumstart
   jnum = jnumstart
   Do
      pmff = pmff * (knum * (cs - colsum(c) - rowsum + knum))
      knum = knum - 1#
      pmff = pmff / ((colsum(c) - knum) * (rowsum - knum))
      If pmff = 0# Then Exit Do
      prob_d = fet_25(c - 1, colsum, rowsum - knum, pmf_Obs / pmff, inum, jnum)
      If inum > jnum Then
         prob = prob + cdf_hypergeometric(knum, colsum(c), rowsum, cs)
         Exit Do
      End If
      prob = prob + pmff * prob_d
   Loop
   fet_25 = prob
Else
   fet_25 = "c must be >= 5"
End If
End Function

Function fet_26(c As Long, ByRef colsum() As Double, ByVal rowsum As Double, ByVal pmf_Obs As Double, ByRef inumstart As Double, ByRef jnumstart As Double) As Double
Dim d As Double, cs As Double, colsum1_2 As Double, prob As Double, temp As Double, inum As Double, jnum As Double, inum_save As Double, jnum_save As Double
Dim cdf As Double, ccdf As Double, pmf_table_inum As Double, pmf_table_jnum As Double, mode As Double, col1mRowSum As Double, row3sum As Double
Dim i As Long, j As Long, tc As Long
Dim all_d_zero As Boolean

Dim d4num As Double, d4num_down As Double, d4num_up As Double
Dim pmff As Double, pmfd4 As Double, pmfd4_down As Double, pmfd4_up As Double
Dim cdff As Double, probf As Double, pmf_Obs_save As Double, cdf_start As Double, ccdf_start As Double
Dim fnum As Double, fnum_start As Double, pmff_start As Double, rowsummfnum As Double
Dim cdf_save As Double, ccdf_save As Double, inum_min As Double, jnum_max As Double, pmf_table_inum_save As Double, pmf_table_jnum_save As Double

Dim continue_up As Boolean, continue_down As Boolean, exit_loop As Boolean
Dim el5 As Boolean

ReDim ml(1 To c) As Double, dnum(5 To c) As Double, dnum_up(5 To c) As Double, dnum_down(5 To c) As Double, dnum_save(5 To c) As Double, colsumsum(3 To c) As Double
ReDim pmfd(5 To c) As Double, pmfd_save(5 To c) As Double, probf5(5 To c) As Double
ReDim inum_save5(5 To c) As Double, jnum_save5(5 To c) As Double, pmf_table_inum_save5(5 To c) As Double, pmf_table_jnum_save5(5 To c) As Double, cdf_save5(5 To c) As Double, ccdf_save5(5 To c) As Double
ReDim c_up(5 To c) As Boolean, c_down(5 To c) As Boolean
ReDim dnumsum(4 To c) As Double, pmf_prod(4 To c + 1)
ReDim inum_next(5 To c) As Double, jnum_next(5 To c) As Double, pmf_table_inum_next(5 To c) As Double, pmf_table_jnum_next(5 To c) As Double, cdf_next(5 To c) As Double, ccdf_next(5 To c) As Double

If pmf_Obs > 1# Then 'All tables have pmf <= 1
   fet_26 = 1#
   Exit Function
End If
pmf_Obs_save = pmf_Obs

colsumsum(3) = colsum(3)
For i = 4 To c
   colsumsum(i) = colsumsum(i - 1) + colsum(i)
Next i
colsum1_2 = colsum(1) + colsum(2)
cs = colsum1_2 + colsumsum(c)

'First guess at mode vector
ml(1) = rowsum
For i = 2 To c
   ml(i) = Int(rowsum * colsum(i) / cs + 0.5)
   ml(1) = ml(1) - ml(i)
Next i

Do 'Update guess at mode vector
   all_d_zero = True
   For i = 1 To c - 1
      For j = i + 1 To c
         d = ml(i) - Int((colsum(i) + 1#) * (ml(i) + ml(j) + 1#) / (colsum(i) + colsum(j) + 2#))
         If d <> 0# Then
            ml(i) = ml(i) - d
            ml(j) = ml(j) + d
            all_d_zero = False
         End If
      Next j
   Next i
Loop Until all_d_zero

'If c > = 5 And colsum(6) > 1 Then
If c >= 5 Then
   d4num = ml(4)
   dnumsum(c) = 0#
   For i = c To 5 Step -1
      probf5(i) = 0#
      dnum(i) = ml(i)
      dnumsum(i - 1) = dnumsum(i) + dnum(i)
   Next i
   fnum = ml(3) + d4num + dnumsum(4)
   rowsummfnum = rowsum - fnum
   pmff = pmf_hypergeometric(fnum, colsumsum(c), rowsum, cs)
   pmff_start = pmff
   pmf_prod(c + 1) = pmff
   For i = c To 5 Step -1
      pmfd(i) = pmf_hypergeometric(dnum(i), fnum - dnumsum(i), colsum(i), colsumsum(i))
      pmf_prod(i) = pmf_prod(i + 1) * pmfd(i)
   Next i
   pmfd4 = pmf_hypergeometric(d4num, fnum - dnumsum(4), colsum(4), colsumsum(4))
   pmf_Obs = pmf_Obs_save / (pmfd4 * pmf_prod(5))
   col1mRowSum = colsum(1) - rowsum
   inum_min = Max(0#, -(fnum + col1mRowSum))
   inum = Min(Max(inum_min, inumstart), ml(2))
   jnum_max = Min(rowsummfnum, colsum(2))
   jnum = Max(Min(jnum_max, jnumstart), ml(2))
   pmf_table_inum = pmf_hypergeometric(inum, rowsummfnum, colsum(2), colsum1_2)
   mode = Int((rowsummfnum + 1#) * (colsum(2) + 1#) / (colsum1_2 + 2#))
   Do While pmf_table_inum = 0#
      inum = Int((inum + Max(2, mode)) * 0.5)
      pmf_table_inum = pmf_hypergeometric(inum, rowsummfnum, colsum(2), colsum1_2)
   Loop
   Do While (pmf_table_inum > pmf_Obs) And (inum > inum_min)
      pmf_table_inum = pmf_table_inum * (inum * (col1mRowSum + fnum + inum))
      inum = inum - 1#
      pmf_table_inum = pmf_table_inum / ((rowsummfnum - inum) * (colsum(2) - inum))
   Loop
   Do While (pmf_table_inum <= pmf_Obs) And (inum <= mode)
      pmf_table_inum = pmf_table_inum * (rowsummfnum - inum) * (colsum(2) - inum)
      inum = inum + 1#
      pmf_table_inum = pmf_table_inum / (inum * (col1mRowSum + fnum + inum))
   Loop
   pmf_table_jnum = pmf_hypergeometric(jnum, rowsummfnum, colsum(2), colsum1_2)
   Do While pmf_table_jnum = 0#
      jnum = Int((jnum + mode) * 0.5)
      pmf_table_jnum = pmf_hypergeometric(jnum, rowsummfnum, colsum(2), colsum1_2)
   Loop
   Do While (pmf_table_jnum > pmf_Obs) And (jnum < jnum_max)
      pmf_table_jnum = pmf_table_jnum * ((rowsummfnum - jnum) * (colsum(2) - jnum))
      jnum = jnum + 1#
      pmf_table_jnum = pmf_table_jnum / (jnum * (col1mRowSum + fnum + jnum))
   Loop
   Do While (pmf_table_jnum < pmf_Obs) And (jnum >= mode)
      pmf_table_jnum = pmf_table_jnum * (jnum * (col1mRowSum + fnum + jnum))
      jnum = jnum - 1#
      pmf_table_jnum = pmf_table_jnum / ((rowsummfnum - jnum) * (colsum(2) - jnum))
   Loop
'Debug.Print fnum, d4num, dnum(5), dnum(6), inum, jnum, pmf_Obs
   inumstart = inum
   jnumstart = jnum
   If inum > jnum Then
      fet_26 = 1#
      Exit Function
   End If
   cdf = cdf_hypergeometric(inum - 1, rowsummfnum, colsum(2), colsum1_2)
   ccdf = comp_cdf_hypergeometric(jnum, rowsummfnum, colsum(2), colsum1_2)
   fnum_start = fnum
   continue_up = True
   continue_down = True
   inum_save = inum
   jnum_save = jnum
   pmf_table_inum_save = pmf_table_inum
   pmf_table_jnum_save = pmf_table_jnum
   cdf_save = cdf
   ccdf_save = ccdf
   cdf_start = cdf
   ccdf_start = ccdf
   prob = 0#
   Do
      For i = 5 To c
         inum_next(i) = inum
         jnum_next(i) = jnum
         pmf_table_inum_next(i) = pmf_table_inum
         pmf_table_jnum_next(i) = pmf_table_jnum
         cdf_next(i) = cdf
         ccdf_next(i) = ccdf
         dnum_save(i) = dnum(i)
         pmfd_save(i) = pmfd(i)
         inum_save5(i) = inum
         jnum_save5(i) = jnum
         pmf_table_inum_save5(i) = pmf_table_inum
         pmf_table_jnum_save5(i) = pmf_table_jnum
         cdf_save5(i) = cdf
         ccdf_save5(i) = ccdf
         c_up(i) = True
         c_down(i) = True
      Next i
      Do
         probf = pmfd4 * (cdf + ccdf)
         pmf_Obs = pmf_Obs * pmfd4 ' was pmf_Obs = pmf_Obs / (pmfd4 * pmf_prod(5))
         pmfd4_down = pmfd4
         pmfd4_up = pmfd4
         d4num_down = d4num
         d4num_up = d4num
         pmfd4_down = pmfd4_down * d4num_down * (colsum(3) - fnum + dnumsum(4) + d4num_down)
         d4num_down = d4num_down - 1
         pmfd4_down = pmfd4_down / ((fnum - dnumsum(4) - d4num_down) * (colsum(4) - d4num_down))
         pmfd4_up = pmfd4_up * ((fnum - dnumsum(4) - d4num_up) * (colsum(4) - d4num_up))
         d4num_up = d4num_up + 1#
         pmfd4_up = pmfd4_up / (d4num_up * (colsum(3) - fnum + dnumsum(4) + d4num_up))
         inum = inum
         Do
            pmfd4 = Max(pmfd4_down, pmfd4_up)
            If pmfd4 = 0# Then Exit Do
            If pmfd4 * pmf_table_inum <= pmf_Obs Then
               Do
                  cdf = cdf + pmf_table_inum
                  pmf_table_inum = pmf_table_inum * (rowsummfnum - inum) * (colsum(2) - inum)
                  inum = inum + 1#
                  pmf_table_inum = pmf_table_inum / (inum * (col1mRowSum + fnum + inum))
               Loop Until (pmfd4 * pmf_table_inum > pmf_Obs) Or (inum > mode)
            End If
            If pmfd4 * pmf_table_jnum < pmf_Obs Then
               Do
                  ccdf = ccdf + pmf_table_jnum
                  pmf_table_jnum = pmf_table_jnum * (jnum * (col1mRowSum + fnum + jnum))
                  jnum = jnum - 1#
                  pmf_table_jnum = pmf_table_jnum / ((rowsummfnum - jnum) * (colsum(2) - jnum))
               Loop Until (pmfd4 * pmf_table_jnum > pmf_Obs) Or (jnum < mode)
            End If
            If inum > jnum Then
               pmfd4_down = cdf_hypergeometric(d4num_down, fnum - dnumsum(4), colsum(4), colsumsum(4))
               pmfd4_up = comp_cdf_hypergeometric(d4num_up - 1#, fnum - dnumsum(4), colsum(4), colsumsum(4))
               probf = probf + pmfd4_down + pmfd4_up
               Exit Do
            End If
            probf = probf + pmfd4 * (cdf + ccdf)
            If pmfd4_down > pmfd4_up Then
               pmfd4_down = pmfd4_down * d4num_down * (colsum(3) - fnum + dnumsum(4) + d4num_down)
               d4num_down = d4num_down - 1#
               pmfd4_down = pmfd4_down / ((fnum - dnumsum(4) - d4num_down) * (colsum(4) - d4num_down))
            Else
               pmfd4_up = pmfd4_up * ((fnum - dnumsum(4) - d4num_up) * (colsum(4) - d4num_up))
               d4num_up = d4num_up + 1#
               pmfd4_up = pmfd4_up / (d4num_up * (colsum(3) - fnum + dnumsum(4) + d4num_up))
            End If
'Debug.Print inum, jnum, pmf_table_inum, pmf_table_jnum, cdf, ccdf, probf
         Loop
         tc = 5
         probf5(tc) = probf5(tc) + probf * pmfd(tc)
         Do
            Do
               el5 = True
               If c_up(tc) Then
                  pmfd(tc) = pmfd(tc) * ((fnum - dnumsum(tc - 1)) * (colsum(tc) - dnum(tc)))
                  dnum(tc) = dnum(tc) + 1#
                  dnumsum(tc - 1) = dnumsum(tc) + dnum(tc)
                  pmfd(tc) = pmfd(tc) / (dnum(tc) * (colsumsum(tc - 1) - fnum + dnumsum(tc - 1)))
                  c_up(tc) = pmfd(tc) > 0#
                  If Not c_up(tc) Then
                     pmfd(tc) = pmfd_save(tc)
                     dnum(tc) = dnum_save(tc)
                     dnumsum(tc - 1) = dnumsum(tc) + dnum(tc)
                     inum_next(tc) = inum_save5(tc)
                     jnum_next(tc) = jnum_save5(tc)
                     pmf_table_inum_next(tc) = pmf_table_inum_save5(tc)
                     pmf_table_jnum_next(tc) = pmf_table_jnum_save5(tc)
                     cdf_next(tc) = cdf_save5(tc)
                     ccdf_next(tc) = ccdf_save5(tc)
                  End If
               End If
               If Not c_up(tc) Then
                  pmfd(tc) = pmfd(tc) * dnum(tc) * (colsumsum(tc - 1) - fnum + dnumsum(tc - 1))
                  dnum(tc) = dnum(tc) - 1#
                  dnumsum(tc - 1) = dnumsum(tc) + dnum(tc)
                  pmfd(tc) = pmfd(tc) / ((fnum - dnumsum(tc - 1)) * (colsum(tc) - dnum(tc)))
                  c_down(tc) = pmfd(tc) > 0#
                  If Not c_down(tc) Then
                     tc = tc + 1
                     If tc > c Then Exit Do
                     probf5(tc) = probf5(tc) + probf5(tc - 1) * pmfd(tc)
'Debug.Print tc, probf5(tc), dnum(tc), probf5(tc - 1), pmfd(tc)
                     probf5(tc - 1) = 0#
                     el5 = False
                  End If
               End If
            Loop Until el5
            If tc > c Then Exit Do
            
            row3sum = fnum - dnumsum(tc - 1)
            ml(3) = row3sum
            For i = 4 To tc - 1
               ml(i) = Int(row3sum * colsum(i) / colsumsum(tc - 1) + 0.5)
               ml(3) = ml(3) - ml(i)
            Next i
            
            Do 'Update guess at mode vector
               all_d_zero = True
               For i = 3 To tc - 2
                  For j = i + 1 To tc - 1
                     d = ml(i) - Int((colsum(i) + 1#) * (ml(i) + ml(j) + 1#) / (colsum(i) + colsum(j) + 2#))
                     If d <> 0# Then
                        ml(i) = ml(i) - d
                        ml(j) = ml(j) + d
                        all_d_zero = False
                     End If
                  Next j
               Next i
            Loop Until all_d_zero
            pmf_prod(tc) = pmf_prod(tc + 1) * pmfd(tc)
            For i = tc - 1 To 5 Step -1
               dnum(i) = ml(i)
               dnumsum(i - 1) = dnumsum(i) + dnum(i)
               pmfd(i) = pmf_hypergeometric(dnum(i), fnum - dnumsum(i), colsum(i), colsumsum(i))
               pmfd_save(i) = pmfd(i)
               pmf_prod(i) = pmf_prod(i + 1) * pmfd(i)
               dnum_save(i) = dnum(i)
            Next i
            d4num = ml(4)
            dnumsum(4) = dnumsum(5) + dnum(5)
            pmfd4 = pmf_hypergeometric(d4num, fnum - dnumsum(4), colsum(4), colsumsum(4))
            pmf_Obs = pmf_Obs_save / (pmfd4 * pmf_prod(5))
            inum = inum_next(tc)
            jnum = jnum_next(tc)
            pmf_table_inum = pmf_table_inum_next(tc)
            pmf_table_jnum = pmf_table_jnum_next(tc)
            cdf = cdf_next(tc)
            ccdf = ccdf_next(tc)
'Debug.Print fnum, d4num, dnum(5), dnum(6), inum, jnum, pmf_Obs, pmff, pmfd4, pmfd(5), pmfd(6)
            If pmf_table_inum * (inum * (col1mRowSum + fnum + inum)) / ((rowsummfnum - inum + 1#) * (colsum(2) - inum + 1#)) > pmf_Obs Then
               fet_26 = "Problem with cdf"
            End If
            Do While (pmf_table_inum <= pmf_Obs) And (inum <= mode)
               cdf = cdf + pmf_table_inum
               pmf_table_inum = pmf_table_inum * (rowsummfnum - inum) * (colsum(2) - inum)
               inum = inum + 1#
               pmf_table_inum = pmf_table_inum / (inum * (col1mRowSum + fnum + inum))
            Loop
            If pmf_table_jnum * ((rowsummfnum - jnum) * (colsum(2) - jnum)) / ((jnum + 1#) * (col1mRowSum + fnum + jnum + 1#)) > pmf_Obs Then
               fet_26 = "Problem with ccdf"
            End If
            Do While (pmf_table_jnum < pmf_Obs) And (jnum >= mode)
               ccdf = ccdf + pmf_table_jnum
               pmf_table_jnum = pmf_table_jnum * (jnum * (col1mRowSum + fnum + jnum))
               jnum = jnum - 1#
               pmf_table_jnum = pmf_table_jnum / ((rowsummfnum - jnum) * (colsum(2) - jnum))
            Loop
            'el5 = True
            If inum > jnum Then
               If c_up(tc) Then
                  probf5(tc) = probf5(tc) + comp_cdf_hypergeometric(dnum(tc) - 1#, fnum - dnumsum(tc), colsum(tc), colsumsum(tc))
                  pmfd(tc) = pmfd_save(tc)
                  dnum(tc) = dnum_save(tc)
                  dnumsum(tc - 1) = dnumsum(tc) + dnum(tc)
                  inum_next(tc) = inum_save5(tc)
                  jnum_next(tc) = jnum_save5(tc)
                  pmf_table_inum_next(tc) = pmf_table_inum_save5(tc)
                  pmf_table_jnum_next(tc) = pmf_table_jnum_save5(tc)
                  cdf_next(tc) = cdf_save5(tc)
                  ccdf_next(tc) = ccdf_save5(tc)
                  c_up(tc) = False
               Else
                  probf5(tc) = probf5(tc) + cdf_hypergeometric(dnum(tc), fnum - dnumsum(tc), colsum(tc), colsumsum(tc))
                  c_down(tc) = False
                  tc = tc + 1
                  If tc > c Then Exit Do
                  probf5(tc) = probf5(tc) + probf5(tc - 1) * pmfd(tc)
'Debug.Print tc, probf5(tc), inum, jnum, dnum(tc), probf5(tc - 1), pmfd(tc)
                  probf5(tc - 1) = 0#
               End If
               el5 = False
            End If
         Loop Until el5
         If tc > c Then Exit Do
         For i = tc To 5 Step -1
            inum_next(i) = inum
            jnum_next(i) = jnum
            pmf_table_inum_next(i) = pmf_table_inum
            pmf_table_jnum_next(i) = pmf_table_jnum
            cdf_next(i) = cdf
            ccdf_next(i) = ccdf
         Next i
         For i = tc - 1 To 5 Step -1
            c_up(i) = True
            c_down(i) = True
            inum_save5(i) = inum
            jnum_save5(i) = jnum
            pmf_table_inum_save5(i) = pmf_table_inum
            pmf_table_jnum_save5(i) = pmf_table_jnum
            cdf_save5(i) = cdf
            ccdf_save5(i) = ccdf
         Next i
      Loop

      prob = prob + probf5(c) * pmff
'Debug.Print c + 1, prob, fnum, probf5(c), pmff
      probf5(c) = 0#
      If continue_up Then
         pmff = pmff * ((rowsummfnum) * (colsumsum(c) - fnum))
         fnum = fnum + 1#
         rowsummfnum = rowsum - fnum
         pmff = pmff / ((colsum1_2 - rowsummfnum) * (fnum))
         continue_up = pmff > 0#
         If Not continue_up Then
            pmff = pmff_start
            fnum = fnum_start
            rowsummfnum = rowsum - fnum
            inum_save = inumstart
            jnum_save = jnumstart
            cdf_save = cdf_start
            ccdf_save = ccdf_start
         Else
            inum = inum_save
            jnum = jnum_save
            cdf = cdf_save
            ccdf = ccdf_save
            inum = inum - 1#
            temp = PBB(col1mRowSum + fnum + inum, rowsummfnum - inum, colsum(2) - inum, inum + 1#)
            hTerm = hTerm * ((rowsummfnum - inum) * (colsum(2) - inum) * (colsum1_2 + 1#))
            inum = inum + 1#
            hTerm = hTerm / ((rowsummfnum + 1#) * (colsum(2) + 1#) * (col1mRowSum + fnum + inum))
            pmf_table_inum = hTerm 'pmf_hypergeometric(inum, rowsummfnum, colsum(2), colsum1_2)
            cdf = cdf + temp
            temp = PBB(jnum, colsum(2) - jnum, rowsummfnum - jnum + 1#, col1mRowSum + fnum + jnum)
            hTerm = hTerm * (jnum * (col1mRowSum + fnum + jnum) * (colsum1_2 + 1#))
            jnum = jnum - 1#
            hTerm = hTerm / ((rowsummfnum + 1#) * (colsum(2) - jnum) * (colsum(1) + 1#))
            pmf_table_jnum = hTerm 'pmf_hypergeometric(jnum, rowsummfnum, colsum(2), colsum1_2)
            ccdf = ccdf + temp
         End If
      End If
      Do
         If Not continue_up Then
            pmff = pmff * ((colsum1_2 - rowsummfnum) * (fnum))
            fnum = fnum - 1#
            rowsummfnum = rowsum - fnum
            pmff = pmff / ((rowsummfnum) * (colsumsum(c) - fnum))
            continue_down = pmff > 0#
            If Not continue_down Then Exit Do
            inum = inum_save
            jnum = jnum_save
            cdf = cdf_save
            ccdf = ccdf_save
            temp = PBB(inum, colsum(2) - inum, rowsummfnum - inum, col1mRowSum + fnum + inum + 1#)
            hTerm = hTerm * (rowsummfnum - inum) * (colsum(2) - inum) * (colsum1_2 + 1#)
            inum = inum + 1#
            hTerm = hTerm / (inum * (colsum(1) + 1#) * (colsum1_2 - rowsummfnum + 1#))
            pmf_table_inum = hTerm 'pmf_hypergeometric(inum, rowsummfnum, colsum(2), colsum1_2)
            cdf = cdf + temp
            temp = PBB(col1mRowSum + fnum + jnum + 1#, rowsummfnum - jnum - 1#, colsum(2) - jnum, jnum + 1#)
            hTerm = hTerm * ((jnum + 1#) * (col1mRowSum + fnum + jnum + 1#) * (colsum1_2 + 1#)) / ((colsum(2) + 1#) * (rowsummfnum - jnum) * (colsum1_2 - rowsummfnum + 1#))
            pmf_table_jnum = hTerm 'pmf_hypergeometric(jnum, rowsummfnum, colsum(2), colsum1_2)
            ccdf = ccdf + temp
         End If
         
         For i = c To 5 Step -1
            c_up(i) = True
            c_down(i) = True
            ml(i) = Int(fnum * colsum(i) / colsumsum(c) + 0.5)
            dnumsum(i - 1) = dnumsum(i) + ml(i)
         Next i
         ml(4) = Int(fnum * colsum(4) / colsumsum(c) + 0.5)
         ml(3) = fnum - ml(4) - dnumsum(4)
       
         Do 'Update guess at mode vector
            all_d_zero = True
            For i = 3 To c - 1
               For j = i + 1 To c
                  d = ml(i) - Int((colsum(i) + 1#) * (ml(i) + ml(j) + 1#) / (colsum(i) + colsum(j) + 2#))
                  If d <> 0# Then
                     ml(i) = ml(i) - d
                     ml(j) = ml(j) + d
                     all_d_zero = False
                  End If
               Next j
            Next i
         Loop Until all_d_zero
'Debug.Print ml(3), ml(4), ml(5), ml(6)
         pmf_prod(c + 1) = pmff
         For i = c To 5 Step -1
            dnum(i) = ml(i)
            dnumsum(i - 1) = dnumsum(i) + dnum(i)
            pmfd(i) = pmf_hypergeometric(dnum(i), fnum - dnumsum(i), colsum(i), colsumsum(i))
            pmfd_save(i) = pmfd(i)
            pmf_prod(i) = pmf_prod(i + 1) * pmfd(i)
            dnum_save(i) = dnum(i)
         Next i
         d4num = ml(4)
         mode = Int((rowsummfnum + 1#) * (colsum(2) + 1#) / (colsum1_2 + 2#))
         'If ml(4) <> Int((fnum - dnumsum(4) + 1#) * (colsum(4) + 1#) / (colsumsum(4) + 2#)) Then
         '   fet_26 = "Problem with mode"
         'End If
         pmfd4 = pmf_hypergeometric(d4num, fnum - dnumsum(4), colsum(4), colsumsum(4))
         pmf_Obs = pmf_Obs_save / (pmfd4 * pmf_prod(5))
         inum_min = Max(0#, -(fnum + col1mRowSum))
         jnum_max = Min(rowsummfnum, colsum(2))
         If jnum_max > 0# And pmf_Obs < 1# Then
            Do While (pmf_table_inum > pmf_Obs) And (inum > inum_min) Or (inum > mode)
               pmf_table_inum = pmf_table_inum * (inum * (col1mRowSum + fnum + inum))
               inum = inum - 1#
               pmf_table_inum = pmf_table_inum / ((rowsummfnum - inum) * (colsum(2) - inum))
               cdf = cdf - pmf_table_inum
            Loop
            Do While (pmf_table_inum < pmf_Obs) And (inum <= mode)
               cdf = cdf + pmf_table_inum
               pmf_table_inum = pmf_table_inum * (rowsummfnum - inum) * (colsum(2) - inum)
               inum = inum + 1#
               pmf_table_inum = pmf_table_inum / (inum * (col1mRowSum + fnum + inum))
            Loop
            If inum = inum_min Then cdf = 0#
            Do While (pmf_table_jnum > pmf_Obs) And (jnum < jnum_max) Or (jnum < mode)
               pmf_table_jnum = pmf_table_jnum * ((rowsummfnum - jnum) * (colsum(2) - jnum))
               jnum = jnum + 1#
               pmf_table_jnum = pmf_table_jnum / (jnum * (col1mRowSum + fnum + jnum))
               ccdf = ccdf - pmf_table_jnum
            Loop
            Do While (pmf_table_jnum < pmf_Obs) And (jnum >= mode)
               ccdf = ccdf + pmf_table_jnum
               pmf_table_jnum = pmf_table_jnum * (jnum * (col1mRowSum + fnum + jnum))
               jnum = jnum - 1#
               pmf_table_jnum = pmf_table_jnum / ((rowsummfnum - jnum) * (colsum(2) - jnum))
            Loop
            If jnum = jnum_max Then ccdf = 0#
         Else
            inum = jnum + 1#
         End If
         exit_loop = True
         If inum > jnum Then
            If continue_up Then
               prob = prob + comp_cdf_hypergeometric(fnum - 1#, colsumsum(c), rowsum, cs)
'Debug.Print c + 1, prob, fnum, comp_cdf_hypergeometric(fnum - 1#, colsumsum(c), rowsum, cs)
               continue_up = False
               pmff = pmff_start
               fnum = fnum_start
               rowsummfnum = rowsum - fnum
               inum_save = inumstart
               jnum_save = jnumstart
               cdf_save = cdf_start
               ccdf_save = ccdf_start
               exit_loop = False
            Else
               prob = prob + cdf_hypergeometric(fnum, colsumsum(c), rowsum, cs)
'Debug.Print c + 1, prob, fnum, cdf_hypergeometric(fnum, colsumsum(c), rowsum, cs)
               continue_down = False
            End If
         End If
      Loop Until exit_loop
      If Not continue_down Then Exit Do
'Debug.Print fnum, d4num, dnum(5), dnum(6), inum, jnum, pmf_Obs, pmff, pmfd4, pmfd(5), pmfd(6)
      cdf_save = cdf
      ccdf_save = ccdf
      inum_save = inum
      jnum_save = jnum
      pmf_table_inum_save = pmf_table_inum
      pmf_table_jnum_save = pmf_table_jnum
   Loop Until Not continue_down
   fet_26 = prob
   Exit Function
ElseIf c >= 6 Then
   Dim knum As Double, prob_d As Double
   knum = ml(c)
   pmff = pmf_hypergeometric(knum, colsum(c), rowsum, cs)
   pmff_start = pmff
   inum = inumstart
   jnum = jnumstart
   prob_d = fet_26(c - 1, colsum, rowsum - knum, pmf_Obs / pmff, inum, jnum)
   inumstart = inum
   jnumstart = jnum
   If inum > jnum Then
      fet_26 = 1#
      Exit Function
   End If
   prob = pmff * prob_d

   Do
      pmff = pmff * ((colsum(c) - knum) * (rowsum - knum))
      knum = knum + 1#
      pmff = pmff / (knum * (cs - colsum(c) - rowsum + knum))
      If pmff = 0# Then Exit Do
      prob_d = fet_26(c - 1, colsum, rowsum - knum, pmf_Obs / pmff, inum, jnum)
      If inum > jnum Then
         prob = prob + comp_cdf_hypergeometric(knum - 1#, colsum(c), rowsum, cs)
         Exit Do
      End If
      prob = prob + pmff * prob_d
   Loop
   pmff = pmff_start
   knum = ml(c)
   inum = inumstart
   jnum = jnumstart
   Do
      pmff = pmff * (knum * (cs - colsum(c) - rowsum + knum))
      knum = knum - 1#
      pmff = pmff / ((colsum(c) - knum) * (rowsum - knum))
      If pmff = 0# Then Exit Do
      prob_d = fet_26(c - 1, colsum, rowsum - knum, pmf_Obs / pmff, inum, jnum)
      If inum > jnum Then
         prob = prob + cdf_hypergeometric(knum, colsum(c), rowsum, cs)
         Exit Do
      End If
      prob = prob + pmff * prob_d
   Loop
   fet_26 = prob
Else
   fet_26 = "c must be >= 5"
End If
End Function

Public Function fet(r As Range) As Double
Dim cs As Double, rowsum As Double, rs As Double, maxc As Double, d As Double, a As Double, b As Double, c As Double, bPlusc As Double, cp As Double, rsp As Double, pmf_Obs As Double, pmf_d As Double, pmf_e As Double
Dim pm As Double, cd As Double, prob As Double, pmfh As Double, pmfh_save As Double, temp As Double, inum As Double, jnum As Double, knum As Double, knum_save As Double
Dim rc As Long, cc As Long, i As Long, j As Long, k As Long
Dim inum_save As Double, jnum_save As Double, es11 As Double, es12 As Double, es12_save As Double, es12p13 As Double, es01 As Double, es02 As Double, es023 As Double, pmf_d_save As Double, pmf_e_save As Double, prob_d As Double, mode As Double
Dim all_d_zero As Boolean
rc = r.Rows.count
cc = r.Columns.count
If rc < 2 Or cc < 2 Or Min(rc, cc) >= 3 And Max(rc, cc) >= 4 Then
   fet = [#VALUE!]
   Exit Function
End If
'Change data so that it is 2x3 or 2x4 rather than 3x2 or 4x2.
If rc > cc Then
   i = rc
   rc = cc
   cc = i
   ReDim os(1 To rc, 1 To cc) As Double, es(0 To rc, 0 To cc) As Double, colsum(cc) As Double
   For i = 1 To rc
       For j = 1 To cc
           os(i, j) = r.Item(j, i)
       Next j
   Next i
ElseIf cc * rc = 4 And False Then
   fet = old_fet_22(r.Item(1, 1), r.Item(1, 2), r.Item(2, 1), r.Item(2, 2))
   Exit Function
Else
   ReDim os(1 To rc, 1 To cc) As Double, es(0 To rc, 0 To cc) As Double, colsum(cc) As Double
   For i = 1 To rc
       For j = 1 To cc
           os(i, j) = r.Item(i, j)
       Next j
   Next i
End If
'Calculate row totals and check that all values are non-negative integers
cs = 0#
For i = 1 To rc
   rs = 0#
   For j = 1 To cc
      If os(i, j) < 0 Or Int(os(i, j)) <> os(i, j) Then
         fet = [#VALUE!]
         Exit Function
      End If
      rs = rs + os(i, j)
   Next j
   If rs = 0# Then
      fet = [#VALUE!]
      Exit Function
   End If
   es(i, 0) = rs
   cs = cs + rs
Next i
es(0, 0) = cs
'Calculate column totals and find column with largest total
maxc = 0#
For i = 1 To cc
   rs = 0#
   For j = 1 To rc
      rs = rs + os(j, i)
   Next j
   es(0, i) = rs
   If maxc < rs Then
      k = i
      maxc = rs
   End If
Next i
'Swap largest column into column 1
If k <> 1 Then
   rs = es(0, 1)
   es(0, 1) = maxc
   es(0, k) = rs
   For j = 1 To rc
       rs = os(j, 1)
       os(j, 1) = os(j, k)
       os(j, k) = rs
   Next j
End If
For i = 2 To cc - 1
   maxc = 0#
   For j = i To cc
      If maxc < es(0, j) Then
         k = j
         maxc = es(0, j)
      End If
   Next j
   If k <> i Then
      rs = es(0, i)
      es(0, i) = maxc
      es(0, k) = rs
      For j = 1 To rc
          rs = os(j, i)
          os(j, i) = os(j, k)
          os(j, k) = rs
      Next j
   End If
Next i
If es(0, cc) = 0 Then
   fet = [#VALUE!]
   Exit Function
End If

If rc = 2 Then
   If es(1, 0) < es(2, 0) Then
      For j = 1 To cc
          rs = os(1, j)
          os(1, j) = os(2, j)
          os(2, j) = rs
      Next j
      rs = es(1, 0)
      es(1, 0) = es(2, 0)
      es(2, 0) = rs
   End If
   For j = 0 To cc
      colsum(j) = es(0, j)
   Next j
   rowsum = es(2, 0)
   'Guess at start points for most likely table
   rs = 0#
   For i = 1 To rc
      For j = 1 To cc
         rsp = es(i, 0) * es(0, j) / cs
         rs = rs + (os(i, j) - rsp) ^ 2 / rsp
      Next j
   Next i
   knum = rowsum * colsum(2) / colsum(0)
   If cc > 2 Then
      rsp = cs * (1# / rowsum) * (1# / colsum(cc - 1) + 1# / colsum(cc))
      pm = Sqr(rs / rsp)
   Else
      pm = Abs(os(2, 2) - knum)
   End If
   inum = Int(knum - pm + 0.5)
   jnum = Int(knum + pm + 0.5)
   pmf_Obs = 1#
   For i = cc To 2 Step -1
      pmf_Obs = pmf_Obs * pmf_hypergeometric(os(2, i), colsum(i), rowsum, cs)
      rowsum = rowsum - os(2, i)
      cs = cs - colsum(i)
   Next i
   If pmf_Obs = 0# Then
      fet = 0#
      Exit Function
   End If
   pmf_Obs = pmf_Obs * (1 + 0.000000000000001 * (10 - Log(pmf_Obs)))
   If cc = 2 Then
      fet = fet_22(cc, colsum, es(2, 0), pmf_Obs, inum, jnum)
   ElseIf cc = 3 Then
      fet = fet_23(cc, colsum, es(2, 0), pmf_Obs, inum, jnum)
   ElseIf cc = 4 Then
      fet = fet_24(cc, colsum, es(2, 0), pmf_Obs, inum, jnum)
   ElseIf cc >= 5 Then
      fet = fet_25(cc, colsum, es(2, 0), pmf_Obs, inum, jnum)
   Else
      fet = fet_26(cc, colsum, es(2, 0), pmf_Obs, inum, jnum)
   End If
   Exit Function
ElseIf rc = 3 Then
   If es(1, 0) > es(2, 0) Then
      For j = 1 To cc
          rs = os(1, j)
          os(1, j) = os(2, j)
          os(2, j) = rs
      Next j
      rs = es(1, 0)
      es(1, 0) = es(2, 0)
      es(2, 0) = rs
   End If
   If es(2, 0) > es(3, 0) Then
      For j = 1 To cc
         rs = os(2, j)
         os(2, j) = os(3, j)
         os(3, j) = rs
      Next j
      rs = es(2, 0)
      es(2, 0) = es(3, 0)
      es(3, 0) = rs
      If es(1, 0) > es(2, 0) Then
         For j = 1 To cc
            rs = os(1, j)
            os(1, j) = os(2, j)
            os(2, j) = rs
         Next j
         rs = es(1, 0)
         es(1, 0) = es(2, 0)
         es(2, 0) = rs
      End If
   End If
'If sum of first row > sum of third column then transpose data
   If es(1, 0) > es(0, 3) Then
      rs = os(1, 1)
      os(1, 1) = os(3, 3)
      os(3, 3) = rs
      rs = os(1, 2)
      os(1, 2) = os(2, 3)
      os(2, 3) = rs
      rs = os(2, 1)
      os(2, 1) = os(3, 2)
      os(3, 2) = rs
      rs = es(1, 0)
      es(1, 0) = es(0, 3)
      es(0, 3) = rs
      rs = es(2, 0)
      es(2, 0) = es(0, 2)
      es(0, 2) = rs
      rs = es(3, 0)
      es(3, 0) = es(0, 1)
      es(0, 1) = rs
   End If
ElseIf cc > 3 Then
   fet = [#VALUE!]
   Exit Function
End If

'Initial guess at mode
For i = 1 To cc
   es(rc, i) = es(0, i)
Next i
rs = 0#
For i = 1 To rc
   If i < rc Then es(i, cc) = es(i, 0)
   For j = 1 To cc
      rsp = es(i, 0) * es(0, j) / cs
      rs = rs + (os(i, j) - rsp) ^ 2 / rsp
      If (i < rc) And (j < cc) Then
         es(i, j) = Int(rsp + 0.5)
         es(i, cc) = es(i, cc) - es(i, j)
         es(rc, j) = es(rc, j) - es(i, j)
      End If
   Next j
   If i < rc Then es(rc, cc) = es(rc, cc) - es(i, cc)
Next i

es01 = es(0, 1)
es02 = es(0, 2)
es023 = es02 + es(0, 3)
pmf_d = pmf_hypergeometric(os(1, 1), es01, es(1, 0), cs)
pmf_e = pmf_hypergeometric(os(1, 2), es02, os(1, 2) + os(1, 3), es023)
pmf_Obs = pmf_d * pmf_e * pmf_hypergeometric(os(2, 1), os(2, 1) + os(3, 1), es(2, 0), es(2, 0) + es(3, 0)) * pmf_hypergeometric(os(2, 2), os(2, 2) + os(3, 2), os(2, 2) + os(2, 3), os(2, 2) + os(3, 2) + os(2, 3) + os(3, 3))
pmf_Obs = pmf_Obs * (1.0000000000001)
If pmf_Obs = 0# Then
   fet = 0#
   Exit Function
End If


'Refining guess for mode
Do
   all_d_zero = True
   d = es(2, 2) - Int((es(2, 2) + es(2, 3) + 1#) * (es(2, 2) + es(3, 2) + 1#) / (es(2, 2) + es(2, 3) + es(3, 2) + es(3, 3) + 2#))
   If d <> 0 Then
      es(2, 2) = es(2, 2) - d
      es(2, 3) = es(2, 3) + d
      es(3, 2) = es(3, 2) + d
      es(3, 3) = es(3, 3) - d
      all_d_zero = False
   End If
   d = es(1, 2) - Int((es(1, 2) + es(1, 3) + 1#) * (es(1, 2) + es(3, 2) + 1#) / (es(1, 2) + es(1, 3) + es(3, 2) + es(3, 3) + 2#))
   If d <> 0 Then
      es(1, 2) = es(1, 2) - d
      es(1, 3) = es(1, 3) + d
      es(3, 2) = es(3, 2) + d
      es(3, 3) = es(3, 3) - d
      all_d_zero = False
   End If
   d = es(1, 2) - Int((es(1, 2) + es(1, 3) + 1#) * (es(1, 2) + es(2, 2) + 1#) / (es(1, 2) + es(1, 3) + es(2, 2) + es(2, 3) + 2#))
   If d <> 0 Then
      es(1, 2) = es(1, 2) - d
      es(1, 3) = es(1, 3) + d
      es(2, 2) = es(2, 2) + d
      es(2, 3) = es(2, 3) - d
      all_d_zero = False
   End If
   d = es(2, 1) - Int((es(2, 1) + es(2, 3) + 1#) * (es(2, 1) + es(3, 1) + 1#) / (es(2, 1) + es(2, 3) + es(3, 1) + es(3, 3) + 2#))
   If d <> 0 Then
      es(2, 1) = es(2, 1) - d
      es(2, 3) = es(2, 3) + d
      es(3, 1) = es(3, 1) + d
      es(3, 3) = es(3, 3) - d
      all_d_zero = False
   End If
   d = es(1, 1) - Int((es(1, 1) + es(1, 3) + 1#) * (es(1, 1) + es(3, 1) + 1#) / (es(1, 1) + es(1, 3) + es(3, 1) + es(3, 3) + 2#))
   If d <> 0 Then
      es(1, 1) = es(1, 1) - d
      es(1, 3) = es(1, 3) + d
      es(3, 1) = es(3, 1) + d
      es(3, 3) = es(3, 3) - d
      all_d_zero = False
   End If
   d = es(1, 1) - Int((es(1, 1) + es(1, 3) + 1#) * (es(1, 1) + es(2, 1) + 1#) / (es(1, 1) + es(1, 3) + es(2, 1) + es(2, 3) + 2#))
   If d <> 0 Then
      es(1, 1) = es(1, 1) - d
      es(1, 3) = es(1, 3) + d
      es(2, 1) = es(2, 1) + d
      es(2, 3) = es(2, 3) - d
      all_d_zero = False
   End If
   d = es(2, 1) - Int((es(2, 1) + es(2, 2) + 1#) * (es(2, 1) + es(3, 1) + 1#) / (es(2, 1) + es(2, 2) + es(3, 1) + es(3, 2) + 2#))
   If d <> 0 Then
      es(2, 1) = es(2, 1) - d
      es(2, 2) = es(2, 2) + d
      es(3, 1) = es(3, 1) + d
      es(3, 2) = es(3, 2) - d
      all_d_zero = False
   End If
   d = es(1, 1) - Int((es(1, 1) + es(1, 2) + 1#) * (es(1, 1) + es(3, 1) + 1#) / (es(1, 1) + es(1, 2) + es(3, 1) + es(3, 2) + 2#))
   If d <> 0 Then
      es(1, 1) = es(1, 1) - d
      es(1, 2) = es(1, 2) + d
      es(3, 1) = es(3, 1) + d
      es(3, 2) = es(3, 2) - d
      all_d_zero = False
   End If
   d = es(1, 1) - Int((es(1, 1) + es(1, 2) + 1#) * (es(1, 1) + es(2, 1) + 1#) / (es(1, 1) + es(1, 2) + es(2, 1) + es(2, 2) + 2#))
   If d <> 0 Then
      es(1, 1) = es(1, 1) - d
      es(1, 2) = es(1, 2) + d
      es(2, 1) = es(2, 1) + d
      es(2, 2) = es(2, 2) - d
      all_d_zero = False
   End If
Loop Until all_d_zero
'Guess at start points for most likely table
rsp = cs * (1# / es(2, 0) + 1# / es(3, 0)) * (1# / es(0, 2) + 1# / es(0, 3))
pm = Sqr(rs / rsp)

knum = es(2, 0) * es(0, 2) / es(0, 0)
inum = Int(knum - pm + 0.5)
jnum = Int(knum + pm + 0.5)
inum_save = inum
jnum_save = jnum
'Work from mode out
es12 = es(1, 2)
es11 = es(1, 1)
es12_save = es12
es12p13 = es12 + es(1, 3)
pmf_d = pmf_hypergeometric(es11, es01, es(1, 0), cs)
pmf_e = pmf_hypergeometric(es12, es02, es12p13, es023)
pmf_d_save = pmf_d
pmf_e_save = pmf_e
For j = 0 To rc
   colsum(j) = es(0, j) - es(1, j)
Next j
rowsum = es(2, 0)

prob = 0#
prob_d = fet_23(3, colsum, rowsum, pmf_Obs / (pmf_d * pmf_e), inum, jnum)
If inum > jnum Then
   fet = 1#
   Exit Function
End If
prob_d = prob_d * pmf_e

Do While pmf_d > 0#
   Do
      colsum(2) = colsum(2) - 1#
      colsum(3) = colsum(3) + 1#
      pmf_e = pmf_e * ((es12p13 - es12) * (es02 - es12))
      es12 = es12 + 1#
      pmf_e = pmf_e / (es12 * (es023 - es02 - es12p13 + es12))
      If pmf_e = 0# Then Exit Do
      prob_d = prob_d + pmf_e * fet_23(3, colsum, rowsum, pmf_Obs / (pmf_d * pmf_e), inum, jnum)
      If inum > jnum Then
         prob_d = prob_d + comp_cdf_hypergeometric(es12, es02, es12p13, es023)
         Exit Do
      End If
   Loop
   inum = inum_save
   jnum = jnum_save
   es12 = es12_save
   colsum(2) = es02 - es12
   colsum(3) = colsum(0) - colsum(1) - colsum(2)
   pmf_e = pmf_e_save
   Do
      colsum(2) = colsum(2) + 1#
      colsum(3) = colsum(3) - 1#
      pmf_e = pmf_e * (es12 * (es023 - es02 - es12p13 + es12))
      es12 = es12 - 1#
      pmf_e = pmf_e / ((es12p13 - es12) * (es02 - es12))
      If pmf_e = 0# Then Exit Do
      prob_d = prob_d + pmf_e * fet_23(3, colsum, rowsum, pmf_Obs / (pmf_d * pmf_e), inum, jnum)
      If inum > jnum Then
         prob_d = prob_d + cdf_hypergeometric(es12 - 1#, es02, es12p13, es023)
         Exit Do
      End If
   Loop
   prob = prob + prob_d * pmf_d
   inum = inum_save
   jnum = jnum_save
   pmf_d = pmf_d * ((es01 - es11) * (es(1, 0) - es11))
   es11 = es11 + 1#
   es12p13 = es12p13 - 1#
   pmf_d = pmf_d / (es11 * (cs - es01 - es(1, 0) + es11))
   If pmf_d = 0# Then Exit Do
   es12 = Int((es02 + 1#) * (es12p13 + 1#) / (es023 + 2#))
   colsum(1) = es01 - es11
   colsum(2) = es02 - es12
   colsum(3) = colsum(0) - colsum(1) - colsum(2)
'Guess of mode values for second row
   'Es[1,1] = es11 Don't touch Es[1,1]
   es(1, 2) = es12
   es(1, 3) = es(1, 0) - es11 - es12
   es(2, 1) = Int(rowsum * colsum(1) / colsum(0) + 0.5)
   es(2, 2) = Int(rowsum * colsum(2) / colsum(0) + 0.5)
   es(2, 3) = rowsum - es(2, 1) - es(2, 2)
   es(3, 1) = colsum(1) - es(2, 1)
   es(3, 2) = colsum(2) - es(2, 2)
   es(3, 3) = colsum(3) - es(2, 3)
   
'Refining guess for mode with Es[1,1] fixed.
   Do
      all_d_zero = True
      d = es(2, 2) - Int((es(2, 2) + es(2, 3) + 1#) * (es(2, 2) + es(3, 2) + 1#) / (es(2, 2) + es(2, 3) + es(3, 2) + es(3, 3) + 2#))
      If d <> 0 Then
         es(2, 2) = es(2, 2) - d
         es(2, 3) = es(2, 3) + d
         es(3, 2) = es(3, 2) + d
         es(3, 3) = es(3, 3) - d
         all_d_zero = False
      End If
      d = es(1, 2) - Int((es(1, 2) + es(1, 3) + 1#) * (es(1, 2) + es(3, 2) + 1#) / (es(1, 2) + es(1, 3) + es(3, 2) + es(3, 3) + 2#))
      If d <> 0 Then
         es(1, 2) = es(1, 2) - d
         es(1, 3) = es(1, 3) + d
         es(3, 2) = es(3, 2) + d
         es(3, 3) = es(3, 3) - d
         all_d_zero = False
      End If
      d = es(1, 2) - Int((es(1, 2) + es(1, 3) + 1#) * (es(1, 2) + es(2, 2) + 1#) / (es(1, 2) + es(1, 3) + es(2, 2) + es(2, 3) + 2#))
      If d <> 0 Then
         es(1, 2) = es(1, 2) - d
         es(1, 3) = es(1, 3) + d
         es(2, 2) = es(2, 2) + d
         es(2, 3) = es(2, 3) - d
         all_d_zero = False
      End If
      d = es(2, 1) - Int((es(2, 1) + es(2, 3) + 1#) * (es(2, 1) + es(3, 1) + 1#) / (es(2, 1) + es(2, 3) + es(3, 1) + es(3, 3) + 2#))
      If d <> 0 Then
         es(2, 1) = es(2, 1) - d
         es(2, 3) = es(2, 3) + d
         es(3, 1) = es(3, 1) + d
         es(3, 3) = es(3, 3) - d
         all_d_zero = False
      End If
      d = es(2, 1) - Int((es(2, 1) + es(2, 2) + 1#) * (es(2, 1) + es(3, 1) + 1#) / (es(2, 1) + es(2, 2) + es(3, 1) + es(3, 2) + 2#))
      If d <> 0 Then
         es(2, 1) = es(2, 1) - d
         es(2, 2) = es(2, 2) + d
         es(3, 1) = es(3, 1) + d
         es(3, 2) = es(3, 2) - d
         all_d_zero = False
      End If
   Loop Until all_d_zero
   
   es12 = es(1, 2)
   es12_save = es12
   colsum(2) = es02 - es12
   colsum(3) = colsum(0) - colsum(1) - colsum(2)
   pmf_e = pmf_hypergeometric(es12, es02, es12p13, es023)
   pmf_e_save = pmf_e
   prob_d = fet_23(3, colsum, rowsum, pmf_Obs / (pmf_d * pmf_e), inum, jnum)
   If inum > jnum Then
      prob = prob + comp_cdf_hypergeometric(es11 - 1#, es01, es(1, 0), cs)
      Exit Do
   End If
   prob_d = prob_d * pmf_e
   inum_save = inum
   jnum_save = jnum
Loop
inum = Int(knum - pm + 0.5)
jnum = Int(knum + pm + 0.5)
pmf_d = pmf_d_save
es11 = es(1, 1)
es12 = es(1, 2)
es12p13 = es(1, 0) - es11
Do
   pmf_d = pmf_d * (es11 * (cs - es01 - es(1, 0) + es11))
   es11 = es11 - 1#
   es12p13 = es12p13 + 1#
   pmf_d = pmf_d / ((es01 - es11) * (es(1, 0) - es11))
   If pmf_d = 0# Then Exit Do
   es12 = Int((es02 + 1#) * (es12p13 + 1#) / (es023 + 2#))
   colsum(1) = es01 - es11
   colsum(2) = es02 - es12
   colsum(3) = colsum(0) - colsum(1) - colsum(2)
'Guess of mode values for second row
   'Es[1,1] = es11 Don't touch Es[1,1]
   es(1, 2) = es12
   es(1, 3) = es(1, 0) - es11 - es12
   es(2, 1) = Int(rowsum * colsum(1) / colsum(0) + 0.5)
   es(2, 2) = Int(rowsum * colsum(2) / colsum(0) + 0.5)
   es(2, 3) = rowsum - es(2, 1) - es(2, 2)
   es(3, 1) = colsum(1) - es(2, 1)
   es(3, 2) = colsum(2) - es(2, 2)
   es(3, 3) = colsum(3) - es(2, 3)
   
'Refining guess for mode with Es[1,1] fixed.
   Do
      all_d_zero = True
      d = es(2, 2) - Int((es(2, 2) + es(2, 3) + 1#) * (es(2, 2) + es(3, 2) + 1#) / (es(2, 2) + es(2, 3) + es(3, 2) + es(3, 3) + 2#))
      If d <> 0 Then
         es(2, 2) = es(2, 2) - d
         es(2, 3) = es(2, 3) + d
         es(3, 2) = es(3, 2) + d
         es(3, 3) = es(3, 3) - d
         all_d_zero = False
      End If
      d = es(1, 2) - Int((es(1, 2) + es(1, 3) + 1#) * (es(1, 2) + es(3, 2) + 1#) / (es(1, 2) + es(1, 3) + es(3, 2) + es(3, 3) + 2#))
      If d <> 0 Then
         es(1, 2) = es(1, 2) - d
         es(1, 3) = es(1, 3) + d
         es(3, 2) = es(3, 2) + d
         es(3, 3) = es(3, 3) - d
         all_d_zero = False
      End If
      d = es(1, 2) - Int((es(1, 2) + es(1, 3) + 1#) * (es(1, 2) + es(2, 2) + 1#) / (es(1, 2) + es(1, 3) + es(2, 2) + es(2, 3) + 2#))
      If d <> 0 Then
         es(1, 2) = es(1, 2) - d
         es(1, 3) = es(1, 3) + d
         es(2, 2) = es(2, 2) + d
         es(2, 3) = es(2, 3) - d
         all_d_zero = False
      End If
      d = es(2, 1) - Int((es(2, 1) + es(2, 3) + 1#) * (es(2, 1) + es(3, 1) + 1#) / (es(2, 1) + es(2, 3) + es(3, 1) + es(3, 3) + 2#))
      If d <> 0 Then
         es(2, 1) = es(2, 1) - d
         es(2, 3) = es(2, 3) + d
         es(3, 1) = es(3, 1) + d
         es(3, 3) = es(3, 3) - d
         all_d_zero = False
      End If
      d = es(2, 1) - Int((es(2, 1) + es(2, 2) + 1#) * (es(2, 1) + es(3, 1) + 1#) / (es(2, 1) + es(2, 2) + es(3, 1) + es(3, 2) + 2#))
      If d <> 0 Then
         es(2, 1) = es(2, 1) - d
         es(2, 2) = es(2, 2) + d
         es(3, 1) = es(3, 1) + d
         es(3, 2) = es(3, 2) - d
         all_d_zero = False
      End If
   Loop Until all_d_zero
   
   es12 = es(1, 2)
   es12_save = es12
   colsum(2) = es02 - es12
   colsum(3) = colsum(0) - colsum(1) - colsum(2)
   pmf_e = pmf_hypergeometric(es12, es02, es12p13, es023)
   pmf_e_save = pmf_e
   prob_d = fet_23(3, colsum, rowsum, pmf_Obs / (pmf_d * pmf_e), inum, jnum)
   If inum > jnum Then
      prob = prob + cdf_hypergeometric(es11, es01, es(1, 0), cs)
      Exit Do
   End If
   prob_d = prob_d * pmf_e
   inum_save = inum
   jnum_save = jnum
   
   Do
      colsum(2) = colsum(2) - 1#
      colsum(3) = colsum(3) + 1#
      pmf_e = pmf_e * ((es12p13 - es12) * (es02 - es12))
      es12 = es12 + 1#
      pmf_e = pmf_e / (es12 * (es023 - es02 - es12p13 + es12))
      If pmf_e = 0# Then Exit Do
      prob_d = prob_d + pmf_e * fet_23(3, colsum, rowsum, pmf_Obs / (pmf_d * pmf_e), inum, jnum)
      If inum > jnum Then
         prob_d = prob_d + comp_cdf_hypergeometric(es12, es02, es12p13, es023)
         Exit Do
      End If
   Loop
   inum = inum_save
   jnum = jnum_save
   es12 = es12_save
   pmf_e = pmf_e_save
   colsum(2) = es02 - es12
   colsum(3) = colsum(0) - colsum(1) - colsum(2)
   Do
      colsum(2) = colsum(2) + 1#
      colsum(3) = colsum(3) - 1#
      pmf_e = pmf_e * (es12 * (es023 - es02 - es12p13 + es12))
      es12 = es12 - 1#
      pmf_e = pmf_e / ((es12p13 - es12) * (es02 - es12))
      If pmf_e = 0# Then Exit Do
      prob_d = prob_d + pmf_e * fet_23(3, colsum, rowsum, pmf_Obs / (pmf_d * pmf_e), inum, jnum)
      If inum > jnum Then
         prob_d = prob_d + cdf_hypergeometric(es12 - 1#, es02, es12p13, es023)
         Exit Do
      End If
   Loop
   prob = prob + prob_d * pmf_d
   inum = inum_save
   jnum = jnum_save
Loop

fet = prob
End Function

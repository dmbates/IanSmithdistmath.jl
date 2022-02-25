#// Version 3.4.10                          # requires >= Excel 2000 (for long lines [for Array assignment])
#// Thanks to Jerry W. Lewis for lots of help with testing and improvements to the code.

#// Copyright © [2022] [Ian Smith]

#// Permission is hereby granted, free of charge, to any person obtaining a copy
#// of this software and associated documentation files (the "Software"), to deal
#// in the Software without restriction, including without limitation the rights
#// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
#// copies of the Software, and to permit persons to whom the Software is
#// furnished to do so, subject to the following conditions:

#// The above copyright notice and this permission notice shall be included in all
#// copies or substantial portions of the Software.

#// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
#// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
#// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
#// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
#// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
#// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
#// SOFTWARE.

Option Explicit

Type TValues
   value::Float64
   Log2Adds  As Integer
End Type

Type TAddStack
   Store::Bool
   Where As Integer
   Stack(50) As TValues
End Type

const NonIntegralValuesAllowed_df = true    # Are non-integral degrees of freedom for t, chi_square and f distributions allowed?
const NonIntegralValuesAllowed_NB = true    # Is "successes required" parameter for negative binomial allowed to be non-integral?

const NonIntegralValuesAllowed_Others = true # Can Int function be applied to parameters like sample_size or is it a fault if the parameter is non-integral?

const nc_limit = 1000000.0                   # Upper Limit for non-centrality parameters - as far as I know it#s ok but slower and slower up to 1e12. Above that I don#t know.
const sumAcc = 5e-16
const cfSmall = 0.00000000000001
const cfVSmall = 0.000000000000001
const minLog1Value = -0.79149064
const OneOverSqrTwoPi = 0.39894228040143267793994605993438 # 0.39894228040143267793994605993438
const scalefactor = 1.1579208923731619542357098500869e+77  # 1.1579208923731619542357098500869e+77 = 2^256  # used for rescaling calcs w/o impacting accuracy, to avoid over/underflow
const scalefactor2 = 8.6361685550944446253863518628004e-78 # 8.6361685550944446253863518628004e-78 = 2^-256
const max_discrete = 9007199254740991                      # 2^53 required for exact addition of 1 in hypergeometric routines
const max_crit = 4503599627370496                          # 2^52 to make sure plenty of room for exact addition of 1 in crit routines
const nearly_zero = 9.99999983659714E-317
const cSmall = 5.562684646268003457725581793331e-309       # (smallest number before we start losing precision)/4
const excel0 = 2.2250738585E-308                           # (number smaller than those excel replaces with 0)
const t_nc_limit = 1.3407807929942597e+154                 # just under 1/abs2(cSmall)
const Log1p5 = 0.40546510810816438197801311546435          # 0.40546510810816438197801311546435 = log(1.5)
const logfbit0p5 = 0.054814121051917653896138702348386     # 0.054814121051917653896138702348386 = logfbit(0.5)
const lfb_0 = 8.1061466795327258219670263594382e-02        # 8.1061466795327258219670263594382e-02 = logfbit(0.0)
const lfb_1 = 4.1340695955409294093822081407118e-02        # 4.1340695955409294093822081407118e-02 = logfbit(1.0)
const lfb_2 = 2.7677925684998339148789292746245e-02        # 2.7677925684998339148789292746245e-02 = logfbit(2.0)
const lfb_3 = 2.0790672103765093111522771767849e-02        # 2.0790672103765093111522771767849e-02 = logfbit(3.0)
const lfb_4 = 1.6644691189821192163194865373593e-02        # 1.6644691189821192163194865373593e-02 = logfbit(4.0)
const lfb_5 = 1.3876128823070747998745727023763e-02        # 1.3876128823070747998745727023763e-02 = logfbit(5.0)

#For logfbit functions                       # Stirling#s series for ln(Gamma(x)), A046968/A046969
const lfbc1 = 1.0 / 12.0
const lfbc2 = 1.0 / 30.0                       # lfbc2 on are Sloane#s ratio times 12
const lfbc3 = 1.0 / 105.0
const lfbc4 = 1.0 / 140.0
const lfbc5 = 1.0 / 99.0
const lfbc6 = 691.0 / 30030.0
const lfbc7 = 1.0 / 13.0
const lfbc8 = .35068485511628418514 #.35068485511628418514   # Chosen to make logfbit(6) & logfbit(7) correct
const lfbc9 = 1.6769380337122674863 #1.6769380337122674863   # Chosen to make logfbit(6) & logfbit(7) correct

#For logfbit functions                      #Stieltjes# continued fraction
const cf_0 = 1.0 / 12.0
const cf_1 = 1.0 / 30.0
const cf_2 = 53.0 / 210.0
const cf_3 = 195.0 / 371.0
const cf_4 = 22999.0 / 22737.0
const cf_5 = 29944523.0 / 19733142.0
const cf_6 = 109535241009.0 / 48264275462.0
const cf_7 = 3.0099173832593981700731407342077  #3.0099173832593981700731407342077
const cf_8 = 4.026887192343901226168879531814   #4.026887192343901226168879531814
const cf_9 = 5.0027680807540300516885024122767  #5.0027680807540300516885024122767
const cf_10 = 6.2839113708157821800726631549524 #6.2839113708157821800726631549524
const cf_11 = 7.4959191223840339297523547082674 #7.4959191223840339297523547082674
const cf_12 = 9.0406602343677266995311393604326 #9.0406602343677266995311393604326
const cf_13 = 10.489303654509482277188371304593 #10.489303654509482277188371304593
const cf_14 = 12.297193610386205863989437140092 #12.297193610386205863989437140092
const cf_15 = 13.982876953992430188259760651279 #13.982876953992430188259760651279
const cf_16 = 16.053551416704935469715616365007 #16.053551416704935469715616365007
const cf_17 = 17.976607399870277592569472307671 #17.976607399870277592569472307671
const cf_18 = 20.309762027441653743805414720495 #20.309762027441653743805414720495
const cf_19 = 22.470471639933132495517941571508 #22.470471639933132495517941571508
const cf_20 = 25.065846548945972029163400322506 #25.065846548945972029163400322506
const cf_21 = 27.464451825029133609175558982646 #27.464451825029133609175558982646
const cf_22 = 30.321821231673047126882599306406 #30.321821231673047126882599306406
const cf_23 = 32.958533929972987219994066451412 #32.958533929972987219994066451412
const cf_24 = 36.077698931299242645153320900855 #36.077698931299242645153320900855
const cf_25 = 38.952706682311555734544390410481 #38.952706682311555734544390410481
const cf_26 = 42.333490043576957211381853948856 #42.333490043576957211381853948856
const cf_27 = 45.446960850061621014424175737541 #45.446960850061621014424175737541
const cf_28 = 49.089203129012597708164883350275 #49.089203129012597708164883350275
const cf_29 = 52.441288751415337312569856046996 #52.441288751415337312569856046996
const cf_30 = 56.344845345341843538420365947476 #56.344845345341843538420365947476
const cf_31 = 59.935683907165858207852583492752 #59.935683907165858207852583492752
const cf_32 = 64.100422755920354527906611892238 #64.100422755920354527906611892238
const cf_33 = 67.930140788018221186367702745199 #67.930140788018221186367702745199
const cf_34 = 72.355940555211701969680052963236 #72.355940555211701969680052963236
const cf_35 = 76.424654626829689752585090422288 #76.424654626829689752585090422288
const cf_36 = 81.111403237247965484814230985683 #81.111403237247965484814230985683
const cf_37 = 85.419221276410972614585638717349 #85.419221276410972614585638717349
const cf_38 = 90.366814723864108595513574581683 #90.366814723864108595513574581683
const cf_39 = 94.913837100009887953076231291987 #94.913837100009887953076231291987
const cf_40 = 100.12217846392919748899074683447 #100.12217846392919748899074683447


#For invcnormal                             # http://lib.stat.cmu.edu/apstat/241
const a0 = 3.3871328727963666080            # 3.3871328727963666080
const a1 = 133.14166789178437745            # 133.14166789178437745
const a2 = 1971.5909503065514427            # 1971.5909503065514427
const a3 = 13731.693765509461125            # 13731.693765509461125
const a4 = 45921.953931549871457            # 45921.953931549871457
const a5 = 67265.770927008700853            # 67265.770927008700853
const a6 = 33430.575583588128105            # 33430.575583588128105
const a7 = 2509.0809287301226727            # 2509.0809287301226727
const b1 = 42.313330701600911252            # 42.313330701600911252
const b2 = 687.18700749205790830            # 687.18700749205790830
const b3 = 5394.1960214247511077            # 5394.1960214247511077
const b4 = 21213.794301586595867            # 21213.794301586595867
const b5 = 39307.895800092710610            # 39307.895800092710610
const b6 = 28729.085735721942674            # 28729.085735721942674
const b7 = 5226.4952788528545610            # 5226.4952788528545610
#//Coefficients for P not close to 0, 0.5 or 1.
const c0 = 1.42343711074968357734           # 1.42343711074968357734
const c1 = 4.63033784615654529590           # 4.63033784615654529590
const c2 = 5.76949722146069140550           # 5.76949722146069140550
const c3 = 3.64784832476320460504           # 3.64784832476320460504
const c4 = 1.27045825245236838258           # 1.27045825245236838258
const c5 = 0.241780725177450611770          # 0.241780725177450611770
const c6 = 2.27238449892691845833e-02       # 2.27238449892691845833e-02
const c7 = 7.74545014278341407640e-04       # 7.74545014278341407640e-04
const d1 = 2.05319162663775882187           # 2.05319162663775882187
const d2 = 1.67638483018380384940           # 1.67638483018380384940
const d3 = 0.689767334985100004550          # 0.689767334985100004550
const d4 = 0.148103976427480074590          # 0.148103976427480074590
const d5 = 1.51986665636164571966e-02       # 1.51986665636164571966e-02
const d6 = 5.47593808499534494600e-04       # 5.47593808499534494600e-04
const d7 = 1.05075007164441684324e-09       # 1.05075007164441684324e-09
#//Coefficients for P near 0 or 1.
const e0 = 6.65790464350110377720           # 6.65790464350110377720
const e1 = 5.46378491116411436990           # 5.46378491116411436990
const e2 = 1.78482653991729133580           # 1.78482653991729133580
const e3 = 0.296560571828504891230          # 0.296560571828504891230
const e4 = 2.65321895265761230930e-02       # 2.65321895265761230930e-02
const e5 = 1.24266094738807843860e-03       # 1.24266094738807843860e-03
const e6 = 2.71155556874348757815e-05       # 2.71155556874348757815e-05
const e7 = 2.01033439929228813265e-07       # 2.01033439929228813265e-07
const f1 = 0.599832206555887937690          # 0.599832206555887937690
const f2 = 0.136929880922735805310          # 0.136929880922735805310
const f3 = 1.48753612908506148525e-02       # 1.48753612908506148525e-02
const f4 = 7.86869131145613259100e-04       # 7.86869131145613259100e-04
const f5 = 1.84631831751005468180e-05       # 1.84631831751005468180e-05
const f6 = 1.42151175831644588870e-07       # 1.42151175831644588870e-07
const f7 = 2.04426310338993978564e-15       # 2.04426310338993978564e-15


#For poissapprox                            # Stirling#s series for Gamma(x), A001163/A001164
const coef15 = 1.0 / 12.0
const coef25 = 1.0 / 288.0
const coef35 = -139.0 / 51840.0
const coef45 = -571.0 / 2488320.0
const coef55 = 163879.0 / 209018880.0
const coef65 = 5246819.0 / 75246796800.0
const coef75 = -534703531.0 / 902961561600.0
const coef1 = 2.0 / 3.0                        # Ramanujan#s series for Gamma(x+1,x)-Gamma(x+1)/2, A065973
const coef2 = -4.0 / 135.0                     # cf. http://www.whim.org/nebula/math/gammaratio.html
const coef3 = 8.0 / 2835.0
const coef4 = 16.0 / 8505.0
const coef5 = -8992.0 / 12629925.0
const coef6 = -334144.0 / 492567075.0
const coef7 = 698752.0 / 1477701225.0
const coef8 = 23349012224.0 / 39565450299375.0

const twoThirds = 2.0 / 3.0
const twoFifths = 2.0 / 5.0
const twoSevenths = 2.0 / 7.0
const twoNinths = 2.0 / 9.0
const twoElevenths = 2.0 / 11.0
const twoThirteenths = 2.0 / 13.0

#For binapprox
const oneThird = 1.0 / 3.0
const twoTo27 = 134217728.0                   # 2^27

#For lngammaexpansion
const eulers_const = 0.5772156649015328606065120901 #0.5772156649015328606065120901
const OneMinusEulers_const = 0.4227843350984671393934879099 #0.4227843350984671393934879099
#For logfbit for small args via lngammaexpansion
const HalfMinusEulers_const = -0.0772156649015328606065120901 #-0.0772156649015328606065120901
const Onep25Minusln2Minuseulers_const = -0.020362845461478170023744211558177 #-0.020362845461478170023744211558177
const FiveOver3Minusln3Minuseulers_const = -0.009161286902975885335090660355859 #-0.009161286902975885335090660355859
const Forty7Over48Minusln4Minuseulers_const = -0.00517669268809014610764299968302 #-0.00517669268809014610764299968302
const coeffs0Minusp3125 = 0.00996703342411321823620758332301 #0.00996703342411321823620758332301
const coeffs0Minusp25 = 0.07246703342411321823620758332301 #0.07246703342411321823620758332301
const coeffs0Minus1Third = -0.01086629990922011509333333333333 #-0.01086629990922011509333333333333
const quiteSmall = 0.00000000000001

Dim hTerm::Float64 #Global written to only by PBB to hold the pmf_hypergeometric value.
Dim lfbArray(29)::Float64
Dim lfbArrayInitialised::Bool
Dim coeffs(44)::Float64
Dim coeffsInitialised::Bool
Dim coeffs2(44)::Float64
Dim coeffs2Initialised::Bool

Private Sub InitAddStack(ByRef ast As TAddStack)
ast.Store = true
ast.Where = 0
ast.Stack(0).Log2Adds = 0
ast.Stack(0).value = 0.0
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

Private Sub AddValueToStack(ByRef ast As TAddStack, nextValue::Float64)
if ast.Store
   if ast.Stack(ast.Where).Log2Adds = 0
      ast.Stack(ast.Where).value = nextValue
   else
      ast.Where = ast.Where + 1
      ast.Stack(ast.Where).value = nextValue
      ast.Stack(ast.Where).Log2Adds = 0
   end
else
   ast.Stack(ast.Where).value = ast.Stack(ast.Where).value + nextValue
   ast.Stack(ast.Where).Log2Adds = 1
   while (ast.Where > 0)
      if (ast.Stack(ast.Where).Log2Adds = ast.Stack(ast.Where - 1).Log2Adds)
         ast.Where = ast.Where - 1
         ast.Stack(ast.Where).value = ast.Stack(ast.Where).value + ast.Stack(ast.Where + 1).value
         ast.Stack(ast.Where).Log2Adds = ast.Stack(ast.Where).Log2Adds + 1
      else
         Exit Do
      end
   end
end
ast.Store = !ast.Store
End Sub

function  StackTotal(ByRef ast As TAddStack)::Float64
Dim sum::Float64, c::Float64, t::Float64, y::Float64, i As Integer
   sum = 0.0
   c = 0.0
   For i = ast.Where To 0 Step -1
      y = ast.Stack(i).value - c
      t = sum + y
      c = (t - sum) - y
      sum = t
   Next i
   StackTotal = sum
end

function  TestAddValuesToStack()::Float64
Dim ast As TAddStack
Dim value::Float64
Dim i As Long, k As Long
Call InitAddStack(ast)
value = 1.0000000000001
For k = 1 To 10
   For i = 1 To 100000
      Call AddValueToStack(ast, value)
   Next i
Next k
#value = 10.0
#Call AddValueToStack(ast, value)
#value = 20.0
#Call AddValueToStack(ast, value)
#value = 30.0
#Call AddValueToStack(ast, value)
#value = 40.0
#Call AddValueToStack(ast, value)
#value = 50.0
#Call AddValueToStack(ast, value)
#value = 60.0
#Call AddValueToStack(ast, value)
#value = 70.0
#Call AddValueToStack(ast, value)
#value = 80.0
#Call AddValueToStack(ast, value)
#value = 90.0
#Call AddValueToStack(ast, value)
#value = 100.0
#Call AddValueToStack(ast, value)
TestAddValuesToStack = StackTotal(ast)
end

Private Sub initlfbArray()
if Not lfbArrayInitialised
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
lfbArrayInitialised = true
end
End Sub

Private Sub initCoeffs()
if Not coeffsInitialised
#// for i < UBound coeffs, coeffs[i] holds (zeta(i+2)-1)/(i+2), coeffs[UBound coeffs] holds (zeta(UBound coeffs+2)-1)
coeffs(0)=0.32246703342411321824    #0.32246703342411321824
coeffs(1)=6.7352301053198095133e-02 #6.7352301053198095133e-02
coeffs(2)=2.0580808427784547879e-02 #2.0580808427784547879e-02
coeffs(3)=7.3855510286739852663e-03 #7.3855510286739852663e-03
coeffs(4)=2.8905103307415232858e-03 #2.8905103307415232858e-03
coeffs(5)=1.1927539117032609771e-03 #1.1927539117032609771e-03
coeffs(6)=5.0966952474304242234e-04 #5.0966952474304242234e-04
coeffs(7)=2.2315475845357937976e-04 #2.2315475845357937976e-04
coeffs(8)=9.9457512781808533715e-05 #9.9457512781808533715e-05
coeffs(9)=4.4926236738133141700e-05 #4.4926236738133141700e-05
coeffs(10)=2.0507212775670691553e-05 #2.0507212775670691553e-05
coeffs(11)=9.4394882752683959040e-06 #9.4394882752683959040e-06
coeffs(12)=4.3748667899074878042e-06 #4.3748667899074878042e-06
coeffs(13)=2.0392157538013662368e-06 #2.0392157538013662368e-06
coeffs(14)=9.5514121304074198329e-07 #9.5514121304074198329e-07
coeffs(15)=4.4924691987645660433e-07 #4.4924691987645660433e-07
coeffs(16)=2.1207184805554665869e-07 #2.1207184805554665869e-07
coeffs(17)=1.0043224823968099609e-07 #1.0043224823968099609e-07
coeffs(18)=4.7698101693639805658e-08 #4.7698101693639805658e-08
coeffs(19)=2.2711094608943164910e-08 #2.2711094608943164910e-08
coeffs(20)=1.0838659214896954091e-08 #1.0838659214896954091e-08
coeffs(21)=5.1834750419700466551e-09 #5.1834750419700466551e-09
coeffs(22)=2.4836745438024783172e-09 #2.4836745438024783172e-09
coeffs(23)=1.1921401405860912074e-09 #1.1921401405860912074e-09
coeffs(24)=5.7313672416788620133e-10 #5.7313672416788620133e-10
coeffs(25)=2.7595228851242331452e-10 #2.7595228851242331452e-10
coeffs(26)=1.3304764374244489481e-10 #1.3304764374244489481e-10
coeffs(27)=6.4229645638381000221e-11 #6.4229645638381000221e-11
coeffs(28)=3.1044247747322272762e-11 #3.1044247747322272762e-11
coeffs(29)=1.5021384080754142171e-11 #1.5021384080754142171e-11
coeffs(30)=7.2759744802390796625e-12 #7.2759744802390796625e-12
coeffs(31)=3.5277424765759150836e-12 #3.5277424765759150836e-12
coeffs(32)=1.7119917905596179086e-12 #1.7119917905596179086e-12
coeffs(33)=8.3153858414202848198e-13 #8.3153858414202848198e-13
coeffs(34)=4.0422005252894400655e-13 #4.0422005252894400655e-13
coeffs(35)=1.9664756310966164904e-13 #1.9664756310966164904e-13
coeffs(36)=9.5736303878385557638e-14 #9.5736303878385557638e-14
coeffs(37)=4.6640760264283742246e-14 #4.6640760264283742246e-14
coeffs(38)=2.2737369600659723206e-14 #2.2737369600659723206e-14
coeffs(39)=1.1091399470834522017e-14 #1.1091399470834522017e-14
coeffs(40)=5.4136591567253631315e-15 #5.4136591567253631315e-15
coeffs(41)=2.6438800178609949985e-15 #2.6438800178609949985e-15
coeffs(42)=1.2918959062789967293811764562316e-15 #1.2918959062789967293811764562316e-15
coeffs(43)=6.3159355041984485676779394847024e-16 #6.3159355041984485676779394847024e-16
coeffs(44)=1.421085482803160676983430580383e-14 #1.421085482803160676983430580383e-14
coeffsInitialised = true
end
End Sub
Private Sub initCoeffs2()
if Not coeffs2Initialised
#// coeffs[i] holds (zeta(i+2)-1)/(i+2) - (i+5)/(i+2)/(i+1)*2^(-i-3)
coeffs2(0)=9.96703342411321823620758332301e-3 #9.96703342411321823620758332301e-3
coeffs2(1)=4.85230105319809513323333333333e-3 #4.85230105319809513323333333333e-3
coeffs2(2)=2.35164176111788121233425746862e-3 #2.35164176111788121233425746862e-3
coeffs2(3)=1.1355510286739852662730972914e-3 #1.1355510286739852662730972914e-3
coeffs2(4)=5.4676033074152328575298829848e-4 #5.4676033074152328575298829848e-4
coeffs2(5)=2.6269438789373716759012616902e-4 #2.6269438789373716759012616902e-4
coeffs2(6)=1.2601997117161385090708338501e-4 #1.2601997117161385090708338501e-4
coeffs2(7)=6.03943417869127130948e-5 #6.03943417869127130948e-5
coeffs2(8)=2.89279988929196448257e-5 #2.89279988929196448257e-5
coeffs2(9)=1.38537935563149598820e-5 #1.38537935563149598820e-5
coeffs2(10)=6.63558635521614609862e-6 #6.63558635521614609862e-6
coeffs2(11)=3.17947224962737026300e-6 #3.17947224962737026300e-6
coeffs2(12)=1.52432377823166362836e-6 #1.52432377823166362836e-6
coeffs2(13)=7.31319548444223379639e-7 #7.31319548444223379639e-7
coeffs2(14)=3.51147479316783649952e-7 #3.51147479316783649952e-7
coeffs2(15)=1.68754473874618369035e-7 #1.68754473874618369035e-7
coeffs2(16)=8.11753732546888155551e-8 #8.11753732546888155551e-8
coeffs2(17)=3.90847775936649142159e-8 #3.90847775936649142159e-8
coeffs2(18)=1.88369052760822398681e-8 #1.88369052760822398681e-8
coeffs2(19)=9.08717580313959348175e-9 #9.08717580313959348175e-9
coeffs2(20)=4.38794008336117216467e-9 #4.38794008336117216467e-9
coeffs2(21)=2.12078578473653627963e-9 #2.12078578473653627963e-9
coeffs2(22)=1.02595225309999020577e-9 #1.02595225309999020577e-9
coeffs2(23)=4.96752618206533915776e-10 #4.96752618206533915776e-10
coeffs2(24)=2.40726205228207715756e-10 #2.40726205228207715756e-10
coeffs2(25)=1.16751848407213311847e-10 #1.16751848407213311847e-10
coeffs2(26)=5.66693373586358102002e-11 #5.66693373586358102002e-11
coeffs2(27)=2.75272781658498271913e-11 #2.75272781658498271913e-11
coeffs2(28)=1.33812334011666457419e-11 #1.33812334011666457419e-11
coeffs2(29)=6.50929603319331702812e-12 #6.50929603319331702812e-12
coeffs2(30)=3.16857905287746826547e-12 #3.16857905287746826547e-12
coeffs2(31)=1.54339039998043529180e-12 #1.54339039998043529180e-12
coeffs2(32)=7.52239805801019839357e-13 #7.52239805801019839357e-13
coeffs2(33)=3.66855576849641617566e-13 #3.66855576849641617566e-13
coeffs2(34)=1.79011840661361776213e-13 #1.79011840661361776213e-13
coeffs2(35)=8.73989502840846835258e-14 #8.73989502840846835258e-14
coeffs2(36)=4.26932273880725309600e-14 #4.26932273880725309600e-14
coeffs2(37)=2.08656067727432658676e-14 #2.08656067727432658676e-14
coeffs2(38)=1.02026669800712891581e-14 #1.02026669800712891581e-14
coeffs2(39)=4.99113012967463749398e-15 #4.99113012967463749398e-15
coeffs2(40)=2.44274876330334144841e-15 #2.44274876330334144841e-15
coeffs2(41)=1.19604099925791156327e-15 #1.19604099925791156327e-15
coeffs2(42)=5.85859784064943690566e-16 #5.85859784064943690566e-16
coeffs2(43)=2.87087981566462948467e-16 #2.87087981566462948467e-16
coeffs2(44)=1.40735520163755175636e-16 #1.40735520163755175636e-16
#coeffs2(45)=6.90166402588687226293e-17 #6.90166402588687226293e-17
#coeffs2(46)=3.38578655511664755968e-17 #3.38578655511664755968e-17
#coeffs2(47)=1.66155827667479862403e-17 #1.66155827667479862403e-17
#coeffs2(48)=8.15674061694190990538e-18 #8.15674061694190990539e-18
#coeffs2(49)=4.00551052932053968016e-18 #4.00551052932053968016e-18
#coeffs2(50)=1.96758982791540465538e-18 #1.96758982791540465538e-18
#coeffs2(51)=9.66812504275437439100e-19 #9.66812504275437439100e-19
#coeffs2(52)=4.75200281648237941674e-19 #4.75200281648237941674e-19
#coeffs2(53)=2.33632791481571906191e-19 #2.33632791481571906191e-19
#coeffs2(54)=1.14897269222195236258e-19 #1.14897269222195236258e-19
#coeffs2(55)=5.65198125116716026715e-20 #5.651981251167160267215-20
#coeffs2(56)=2.78101464727381599671e-20 #2.78101464727381599671e-20
#coeffs2(57)=1.36871811383630687470e-20 #1.36871811383630687470e-20
#coeffs2(58)=6.73797960341042295487e-21 #6.73797960341042295487e-21
#coeffs2(59)=3.31777713997525851631e-21 #3.31777713997525851631e-21
#coeffs2(60)=1.63404346465554253504e-21 #1.63404346465554253504e-21
#coeffs2(61)=8.04963210512578619264e-22 #8.04963210512578619264e-22
#coeffs2(62)=3.96626538798230823888e-22 #3.96626538798230823888e-22
#coeffs2(63)=1.95469141675562808544e-22 #1.95469141675562808544e-22
#coeffs2(64)=9.63524657953850607531e-23 #9.63524657953850607531e-23
#coeffs2(65)=4.75043353504700102959e-23 #4.75043353504700102959e-23
#coeffs2(66)=2.34254063551981428148e-23 #2.34254063551981428148e-23
#coeffs2(67)=1.15537315908688659068e-23 #1.15537315908688659068e-23
#coeffs2(68)=5.69949705709986050247e-24 #5.69949705709986050247e-24
#coeffs2(69)=2.81208121322037769865e-24 #2.81208121322037769865e-24
#coeffs2(70)=1.38769580071555495253e-24 #1.38769580071555495253e-24
coeffs2Initialised = true
end
End Sub

function  ec()::Float64
ec = eulers_const
end

function  logcfdersum(x::Float64, i::Float64, d::Float64, Optional derivs As Integer = 0)::Float64
#// Calculation of logcfdersum(x,i,d,derivs) via summation, where derivs is the number of derivatives wrt x required of the normal logcf(x,i,d) function.
#// Will become hopelessly slow for x >= 0.5
Dim tot::Float64, addon::Float64, y::Float64, nd::Float64, k As Integer
tot = 1.0
y = x / (x - 1.0)
nd = (derivs + 1.0) * d
addon = y * nd / (i + nd)
while abs(addon) > abs(0.000000000000001 * tot)
    tot = tot + addon
    nd = nd + d
    addon = addon * y * nd / (i + nd)
end
logcfdersum = (tot + addon) / ((i + derivs * d) * (1.0 - x) ^ (derivs + 1))
For k = 2 To derivs
   logcfdersum = logcfdersum * k
Next k
end

function  deriv2cf(x::Float64, i::Float64, d::Float64)::Float64
#// Accurate calculation of derivative of logcf(x,i,d) wrt x, when x is small in absolute value.
Dim n::Float64, j::Float64, tot::Float64, xtojm1::Float64, addon::Float64
n = i + 2.0 * d
tot = 2.0 / n
j = 2.0
n = n + d
xtojm1 = x
addon = j * (j + 1.0) * xtojm1 / n
while abs(addon) > abs(0.000000000000001 * tot)
    tot = tot + addon
    xtojm1 = x * xtojm1
    j = j + 1.0
    n = n + d
    addon = j * (j + 1.0) * xtojm1 / n
end
deriv2cf = tot + addon
end


function  derivcf(x::Float64, i::Float64, d::Float64)::Float64
#// Accurate calculation of derivative of logcf(x,i,d) wrt x, when x is small in absolute value.
Dim n::Float64, j::Float64, tot::Float64, xtojm1::Float64, addon::Float64
n = i + d
tot = 1.0 / n
j = 2.0
n = n + d
xtojm1 = x
addon = j * xtojm1 / n
while abs(addon) > abs(1E-16 * tot)
    tot = tot + addon
    xtojm1 = x * xtojm1
    j = j + 1.0
    n = n + d
    addon = j * xtojm1 / n
end
derivcf = tot + addon
end

function  min(x::Float64, y::Float64)::Float64
   if x < y
      Min = x
   else
      Min = y
   end
end
function  max(x::Float64, y::Float64)::Float64
   if x > y
      Max = x
   else
      Max = y
   end
end

function  expm1old(x::Float64)::Float64
#// Accurate calculation of exp(x)-1, particularly for small x.
#// Uses a variation of the standard continued fraction for tanh(x) see A&S 4.5.70.
  if (abs(x) < 2)
     Dim a1::Float64, a2::Float64, b1::Float64, b2::Float64, c1::Float64, x2::Float64
     a1 = 24.0
     b1 = 2.0 * (12.0 - x * (6.0 - x))
     x2 = x * x * 0.25
     a2 = 8.0 * (15.0 + x2)
     b2 = 120.0 - x * (60.0 - x * (12.0 - x))
     c1 = 7.0

     while ((abs(a2 * b1 - a1 * b2) > abs(cfSmall * b1 * a2)))

       a1 = c1 * a2 + x2 * a1
       b1 = c1 * b2 + x2 * b1
       c1 = c1 + 2.0

       a2 = c1 * a1 + x2 * a2
       b2 = c1 * b1 + x2 * b2
       c1 = c1 + 2.0
       if (b2 > scalefactor)
         a1 = a1 * scalefactor2
         b1 = b1 * scalefactor2
         a2 = a2 * scalefactor2
         b2 = b2 * scalefactor2
       end
     end

     expm1old = x * a2 / b2
  else
     expm1old = exp(x) - 1.0
  end

end

function  expm1(x::Float64)::Float64
#// Accurate calculation of exp(x)-1, particularly for small x.
#// Based on NR approach to solving log(1+result) = x
Dim y0::Float64, a2::Float64, b1::Float64, b2::Float64, c1::Float64, x2::Float64
  y0 = exp(x) - 1.0
  if abs(x) < 2
     if y0 = 0.0
        expm1 = x
     else
        expm1 = y0 - (log(1.0 + y0) - x) * (1.0 + y0)
     end
  else
     expm1 = y0
  end
end

function  logcf(x::Float64, i::Float64, d::Float64)::Float64
#// Continued fraction for calculation of 1/i + x/(i+d) + x*x/(i+2*d) + x*x*x/(i+3d) + ...
Dim a1::Float64, a2::Float64, b1::Float64, b2::Float64, c1::Float64, c2::Float64, c3::Float64, c4::Float64
     c1 = 2.0 * d
     c2 = i + d
     c4 = c2 + d
     a1 = c2
     b1 = i * (c2 - i * x)
     b2 = d * d * x
     a2 = c4 * c2 - b2
     b2 = c4 * b1 - i * b2

     while ((abs(a2 * b1 - a1 * b2) > abs(cfVSmall * b1 * a2)))

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
       if (b2 > scalefactor)
         a1 = a1 * scalefactor2
         b1 = b1 * scalefactor2
         a2 = a2 * scalefactor2
         b2 = b2 * scalefactor2
       elseif (b2 < scalefactor2)
         a1 = a1 * scalefactor
         b1 = b1 * scalefactor
         a2 = a2 * scalefactor
         b2 = b2 * scalefactor
       end
     end
     logcf = a2 / b2
end

function  logcfplusderiv(x::Float64, i::Float64, d::Float64)::Float64
#// Continued fraction type calculation of derivative of 1/i + x/(i+d) + x*x/(i+2*d) + x*x*x/(i+3d) + ...
Dim a1::Float64, a2::Float64, b1::Float64, b2::Float64, a1dash::Float64, a2dash::Float64, b1dash::Float64, b2dash::Float64, c1::Float64, c2::Float64, c3::Float64, c4::Float64, c5::Float64
     c1 = 2.0 * d
     c2 = i + d
     c4 = c2 + d
     a1 = c2
     a1dash = 0.0
     b1 = i * (c2 - i * x)
     b1dash = -i * i
     b2 = d * d * x
     a2 = c4 * c2 - b2
     a2dash = -d * d
     b2 = c4 * b1 - i * b2
     b2dash = c4 * b1dash + i * a2dash

     while ((abs(a2 * b1 - a1 * b2) > abs(cfVSmall * b1 * a2)))

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
       if (b2 > scalefactor)
         a1 = a1 * scalefactor2
         b1 = b1 * scalefactor2
         a2 = a2 * scalefactor2
         b2 = b2 * scalefactor2
         a1dash = a1dash * scalefactor2
         b1dash = b1dash * scalefactor2
         a2dash = a2dash * scalefactor2
         b2dash = b2dash * scalefactor2
       elseif (b2 < scalefactor2)
         a1 = a1 * scalefactor
         b1 = b1 * scalefactor
         a2 = a2 * scalefactor
         b2 = b2 * scalefactor
         a1dash = a1dash * scalefactor
         b1dash = b1dash * scalefactor
         a2dash = a2dash * scalefactor
         b2dash = b2dash * scalefactor
       end
     end
     c5 = c2 * c2
     c3 = c5 * x
     c4 = c4 + d
     a1dash = c4 * a2dash - c3 * a1dash - c5 * a1
     b1dash = c4 * b2dash - c3 * b1dash - c5 * b1
     a1 = c4 * a2 - c3 * a1
     b1 = c4 * b2 - c3 * b1
     logcfplusderiv = (a1dash * b1 - a1 * b1dash) / b1 ^ 2
end

function  log0Old(x::Float64)::Float64
#//Accurate calculation of log(1+x), particularly for small x.
   Dim term::Float64
   if (abs(x) > 0.5)
      log0Old = log(1.0 + x)
   else
     term = x / (2.0 + x)
     log0Old = 2.0 * term * logcf(term * term, 1.0, 2.0)
   end
end

function  log0(x::Float64)::Float64
#//Accurate and quicker calculation of log(1+x), particularly for small x. Code from Wolfgang Ehrhardt.
   Dim y::Float64
   if x > 4.0
      log0 = log(1.0 + x)
   else
      y = 1.0 + x
      if y = 1.0
         log0 = x
      else
         log0 = log(y) + (x - (y - 1.0)) / y
      end
   end
end

function  lcc(x::Float64)::Float64
#//Accurate calculation of log(1+x)-x, particularly for small x.
   Dim term::Float64, y ::Float64
   if (abs(x) < 0.01)
      term = x / (2.0 + x)
      y = term * term
      lcc = term * ((((2.0 / 9.0 * y + 2.0 / 7.0) * y + 0.4) * y + 2.0 / 3.0) * y - x)
   elseif (x < minLog1Value || x > 1.0)
      lcc = log(1.0 + x) - x
   else
      term = x / (2.0 + x)
      y = term * term
      lcc = term * (2.0 * y * logcf(y, 3.0, 2.0) - x)
   end
end

function  log1(x::Float64)::Float64
#//Accurate calculation of log(1+x)-x, particularly for small x.
   Dim term::Float64, y ::Float64
   if (abs(x) < 0.01)
      term = x / (2.0 + x)
      y = term * term
      log1 = term * ((((2.0 / 9.0 * y + 2.0 / 7.0) * y + 0.4) * y + 2.0 / 3.0) * y - x)
   elseif (x < minLog1Value || x > 1.0)
      log1 = log(1.0 + x) - x
   else
      term = x / (2.0 + x)
      y = term * term
      log1 = term * (2.0 * y * logcf(y, 3.0, 2.0) - x)
   end
end

function  logfbitdif(x::Float64)::Float64
#//Calculation of logfbit(x)-logfbit(1+x). x must be > -1.
  Dim y::Float64, y2::Float64
  if x < -0.65
     logfbitdif = (x + 1.5) * log0(1.0 / (x + 1.0)) - 1.0
  else
     y2 = (2.0 * x + 3.0) ^ -2
     logfbitdif = y2 * logcf(y2, 3.0, 2.0)
  end
end

function  logfbita(x::Float64)::Float64
#//Error part of Stirling#s formula where log(x!) = log(sqrt(twopi))+(x+0.5)*log(x+1)-(x+1)+logfbita(x).
#//Are we ever concerned about the relative error involved in this function? I don#t think so.
  Dim x1::Float64, x2::Float64, x3::Float64
  if (x >= 100000000.0)
     logfbita = lfbc1 / (x + 1.0)
  elseif (x >= 6.0)                      # Abramowitz & Stegun#s series 6.1.41
     x1 = x + 1.0
     x2 = 1.0 / (x1 * x1)
     x3 = x2 * (lfbc6 - x2 * (lfbc7 - x2 * (lfbc8 - x2 * lfbc9)))
     x3 = x2 * (lfbc4 - x2 * (lfbc5 - x3))
     x3 = x2 * (lfbc2 - x2 * (lfbc3 - x3))
     logfbita = lfbc1 * (1.0 - x3) / x1
  elseif (x = 0.0)
     logfbita = lfb_0
  elseif (x = 1.0)
     logfbita = lfb_1
  elseif (x = 2.0)
     logfbita = lfb_2
  elseif (x = 3.0)
     logfbita = lfb_3
  elseif (x = 4.0)
     logfbita = lfb_4
  elseif (x = 5.0)
     logfbita = lfb_5
  elseif (x > -1.0)
     x1 = x
     x2 = 0.0
     while (x1 < 6.0)
        x2 = x2 + logfbitdif(x1)
        x1 = x1 + 1.0
     end
     logfbita = x2 + logfbita(x1)
  else
     logfbita = 1E+308
  end
end

function  logfbitb(x::Float64)::Float64
    Dim lgam::Float64
    Dim i As Integer
    Dim m::Float64
    Dim big::Bool
    Call initCoeffs
    if x <= 0.5
       m = 0.0
    elseif x <= 1.5
       m = 1.0
    elseif x <= 2.5
       m = 2.0
    else
       m = 3.0
    end
    x = x - m
    i = UBound(coeffs)
    lgam = coeffs(i) * logcf(-x / 2.0, i + 2.0, 1.0)
    For i = UBound(coeffs) - 1 To 2 Step -1
       lgam = (coeffs(i) - x * lgam)
    Next i
    if m = 3.0
       lgam = (coeffs(1) - x * lgam)
       logfbitb = ((x * x * (coeffs0Minusp25 - x * lgam) - (x + 3.5) * log1(x / 4.0) + log1(x / 2.0) + log1(x / 3.0)) + Forty7Over48Minusln4Minuseulers_const * x) + lfb_3
    elseif m = 2.0
       lgam = (coeffs(1) - x * lgam)
       logfbitb = ((x * x * (coeffs0Minus1Third - x * lgam) - (x + 2.5) * log1(x / 3.0) + log1(x / 2.0)) + FiveOver3Minusln3Minuseulers_const * x) + lfb_2
    elseif m = 1.0
       #lgam = (coeffs(1) - x * lgam)
       #logfbitb = ((x * x * (coeffs(0) - 0.5 - x * lgam) - (x + 1.5) * log1(x / 2.0)) + Onep25Minusln2Minuseulers_const * x) + lfb_1
       #logfbit = ((coeffs0Minusp3125 - (lgam - 0.125 * (1.0 - (x + 1.5) * logcf(-x / 2.0, 3.0, 1.0))) * x) * x + Onep25Minusln2Minuseulers_const) * x + lfb_1
       #logfbitb = ((coeffs0Minusp3125 - (coeffs(1) - 0.125 - x * lgam + 0.125 * (x + 1.5) * logcf(-x / 2.0, 3.0, 1.0)) * x) * x + Onep25Minusln2Minuseulers_const) * x + lfb_1
       logfbitb = ((coeffs0Minusp3125 - (coeffs(1) - 0.0625 - x * (lgam - 1.0 / 24.0 + 0.0625 * (x + 1.5) * logcf(-x / 2.0, 4.0, 1.0))) * x) * x + Onep25Minusln2Minuseulers_const) * x + lfb_1
    else
       #lgam = (coeffs(1) - x * lgam)
       #logfbitb = ((x * x * (coeffs(0) - 1.0 - x * lgam) - (x + 1.5) * log1(x)) + HalfMinusEulers_const * x) + lfb_0
       #logfbitb = ((coeffs0Minusp25 - ((x + 1.5) * logcf(-x, 3.0, 1.0) - 0.5 + lgam) * x) * x + HalfMinusEulers_const) * x + lfb_0
       #logfbitb = ((coeffs0Minusp25 - (x * (1.0 / 3.0 - (x + 1.5) * logcf(-x, 4.0, 1.0)) + lgam) * x) * x + HalfMinusEulers_const) * x + lfb_0
       logfbitb = ((coeffs0Minusp25 - (coeffs(1) + (x * (x + 1.5) * logcf(-x, 5.0, 1.0) - (6.0 * x + 1.0) / 24.0 - lgam) * x) * x) * x + HalfMinusEulers_const) * x + lfb_0
    end
end

function  logfbit(x::Float64)::Float64
#//Calculates log of x factorial - log(sqrt(2*pi)) +(x+1) -(x+0.5)*log(x+1)
#//using the error part of Stirling#s formula (see Abramowitz & Stegun#s series 6.1.41)
#//and Stieltjes# continued fraction for the gamma function.
#//For x < 1.5, uses expansion of log(x!) and log((x+1)!) from Abramowitz & Stegun#s series 6.1.33
#//We are primarily concerned about the absolute error in this function.
#//Due to cancellation errors in calculating 1+x as x tends to -1, the function loses accuracy and should not be used!
  Dim x1::Float64, x2::Float64, x3::Float64
  if (x >= 6.0)
     x1 = x + 1.0
     if (x >= 1000.0)
        if (x >= 100000000.0)
           x3 = 0.0
        else
           x2 = 1.0 / (x1 * x1)
           x3 = x2 * (lfbc2 - x2 * lfbc3)
        end
     else
        x2 = 1.0 / (x1 * x1)
        if x >= 40.0
           x3 = 0.0
        elseif x >= 15.0
           x3 = x2 * (lfbc6 - x2 * lfbc7)
        else
           x3 = x2 * (lfbc6 - x2 * (lfbc7 - x2 * (lfbc8 - x2 * lfbc9)))
        end
        x3 = x2 * (lfbc4 - x2 * (lfbc5 - x3))
        x3 = x2 * (lfbc2 - x2 * (lfbc3 - x3))
     end
     logfbit = lfbc1 * (1.0 - x3) / x1
     #logfbit = (1.0 - x3) / (12.0 * x1)
  elseif (x = 0.0)
     logfbit = lfb_0
  elseif (x = 1.0)
     logfbit = lfb_1
  elseif (x = 2.0)
     logfbit = lfb_2
  elseif (x = 3.0)
     logfbit = lfb_3
  elseif (x = 4.0)
     logfbit = lfb_4
  elseif (x = 5.0)
     logfbit = lfb_5
  elseif x > 1.5
     x1 = x + 1.0
     if x >= 2.5
        #x2 = 0.25 * ((abs2(x1 * x1 + 81.0) - x1) + 81.0 / (x1 + abs2(x1 * x1 + 90.25)))
        x2 = 40.5 / (x1 + abs2(x1 * x1 + 81.0))
     else
        #x2 = 0.25 * ((abs2(x1 * x1 + 225.0) - x1) + 225.0 / (x1 + abs2(x1 * x1 + 240.25)))
        x2 = 112.5 / (x1 + abs2(x1 * x1 + 225.0))
        x2 = cf_27 / (x1 + cf_28 / (x1 + cf_29 / (x1 + x2)))
        x2 = cf_24 / (x1 + cf_25 / (x1 + cf_26 / (x1 + x2)))
        x2 = cf_21 / (x1 + cf_22 / (x1 + cf_23 / (x1 + x2)))
        x2 = cf_18 / (x1 + cf_19 / (x1 + cf_20 / (x1 + x2)))
     end
     x2 = cf_15 / (x1 + cf_16 / (x1 + cf_17 / (x1 + x2)))
     x2 = cf_12 / (x1 + cf_13 / (x1 + cf_14 / (x1 + x2)))
     x2 = cf_9 / (x1 + cf_10 / (x1 + cf_11 / (x1 + x2)))
     x2 = cf_6 / (x1 + cf_7 / (x1 + cf_8 / (x1 + x2)))
     x2 = cf_3 / (x1 + cf_4 / (x1 + cf_5 / (x1 + x2)))
     #logfbit = cf_0 / (x1 + cf_1 / (x1 + cf_2 / (x1 + x2)))
     logfbit = 1.0 / (12.0 * (x1 + cf_1 / (x1 + cf_2 / (x1 + x2))))
  #elseif (x = 1.5)
  #   logfbit = 3.316287351993628748511050974106e-02   # 3.316287351993628748511050974106e-02
  elseif (x = 0.5)
     logfbit = logfbit0p5                             # 5.481412105191765389613870234839e-02
  elseif (x = -0.5)
     logfbit = 0.15342640972002734529138393927091     # 0.15342640972002734529138393927091
  elseif x >= -0.65
    Dim lgam::Float64
    Dim i As Integer
    if x <= 0.0
       Call initCoeffs
       i = UBound(coeffs)
       lgam = coeffs(i) * logcf(-x / 2.0, i + 2.0, 1.0)
       For i = UBound(coeffs) - 1 To 1 Step -1
          lgam = (coeffs(i) - x * lgam)
       Next i
       logfbit = ((coeffs0Minusp25 - (x * (1.0 / 3.0 - (x + 1.5) * logcf(-x, 4.0, 1.0)) + lgam) * x) * x + HalfMinusEulers_const) * x + lfb_0
    elseif x <= 1.56
       x = x - 1.0
       Call initCoeffs2
       i = UBound(coeffs2) + 3
       lgam = ((x + 2.5) * logcf(-x / 2.0, i, 1.0) - (2.0 / (i - 1.0))) * (2.0 ^ -i) + (3.0 ^ -i) * logcf(-x / 3.0, i, 1.0)
       For i = UBound(coeffs2) To 0 Step -1
          lgam = (coeffs2(i) - x * lgam)
       Next i
       logfbit = (x * lgam + Onep25Minusln2Minuseulers_const) * x + lfb_1
    elseif x <= 2.5
       x = x - 2.0
       Call initCoeffs
       i = UBound(coeffs)
       lgam = coeffs(i) * logcf(-x / 2.0, i + 2.0, 1.0)
       For i = UBound(coeffs) - 1 To 1 Step -1
          lgam = (coeffs(i) - x * lgam)
       Next i
       logfbit = ((x * x * (coeffs0Minus1Third - x * lgam) - (x + 2.5) * log1(x / 3.0) + log1(x / 2.0)) + FiveOver3Minusln3Minuseulers_const * x) + lfb_2
    else
       x = x - 3.0
       Call initCoeffs
       i = UBound(coeffs)
       lgam = coeffs(i) * logcf(-x / 2.0, i + 2.0, 1.0)
       For i = UBound(coeffs) - 1 To 1 Step -1
          lgam = (coeffs(i) - x * lgam)
       Next i
       logfbit = ((x * x * (coeffs0Minusp25 - x * lgam) - (x + 3.5) * log1(x / 4.0) + log1(x / 2.0) + log1(x / 3.0)) + Forty7Over48Minusln4Minuseulers_const * x) + lfb_3
    end
  elseif x > -1.0
    logfbit = logfbitdif(x) + logfbit(x + 1.0)
  else
     logfbit = 1E+308
  end
end

function  lfbaccdif1(a::Float64, b::Float64)::Float64
#//Calculates logfbit(b)-logfbit(a+b) accurately for a > 0 & b >= 0. Reasonably accurate for a >=0 & b < 0.
Dim x1::Float64, x2::Float64, x3::Float64, y1::Float64, y2::Float64, y3::Float64
Dim acc::Float64, i As Integer, Start As Integer, s1::Float64, s2::Float64, tx::Float64, ty::Float64
  if a < 0.0
     lfbaccdif1 = -lfbaccdif1(-a, b + a)
  elseif (b >= 8.0)
     y1 = b + 1.0
     y2 = y1 ^ -2
     x1 = a + b + 1.0
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
     #lfbaccdif1 = lfbc1 * (a * (1.0 - y3) - y1 * acc) / (x1 * y1)
     lfbaccdif1 = (a * (1.0 - y3) - y1 * acc) / (12.0 * x1 * y1)
  elseif b >= 1.7
     y1 = b + 1.0
     x1 = a + b + 1.0
     if b >= 3.0
        Start = 17
     else
        Start = 29
     end
     s1 = (0.5 * (Start + 1.0)) ^ 2
     s2 = (0.5 * (Start + 1.5)) ^ 2
     ty = y1 * abs2(1.0 + s1 * (y1 ^ -2))
     tx = x1 * abs2(1.0 + s1 * (x1 ^ -2))
     y2 = ty - y1
     x2 = tx - x1
     acc = a * (1.0 - (2.0 * y1 + a) / (tx + ty))
     #Seems to work better without the next 2 lines. - !with modification to s2
     ty = y1 * abs2(1.0 + s2 * (y1 ^ -2))
     tx = x1 * abs2(1.0 + s2 * (x1 ^ -2))
     acc = 0.25 * (acc + s1 / ((y1 + ty) * (x1 + tx)) * a * (1.0 + (2.0 * y1 + a) / (tx + ty)))
     y2 = 0.25 * (y2 + s1 / (y1 + ty))
     x2 = 0.25 * (x2 + s1 / (x1 + tx))
     Call initlfbArray
     For i = Start To 1 Step -1
        acc = lfbArray(i) * (a - acc) / ((x1 + x2) * (y1 + y2))
        y2 = lfbArray(i) / (y1 + y2)
        x2 = lfbArray(i) / (x1 + x2)
     Next i
     lfbaccdif1 = cf_0 * (a - acc) / ((x1 + x2) * (y1 + y2))
     #lfbaccdif1 = (a - acc) / (12.0 * (x1 + x2) * (y1 + y2))
  elseif b > -1.0
    Dim scale2::Float64, scale3::Float64
    if b < -0.66
       if a > 1.0
          lfbaccdif1 = logfbitdif(b) + lfbaccdif1(a - 1.0, b + 1.0)
          Exit Function
       elseif a = 1.0
          lfbaccdif1 = logfbitdif(b)
          Exit Function
       else
          s2 = a * log0(1.0 / (b + 1.0 + a))
          s1 = logfbitdif(b + a)
          if s1 > s2
             s1 = (b + 1.5) * log0(a / ((b + 1.0) * (b + 2.0 + a))) - s2
          else
             s2 = s1
             s1 = (logfbitdif(b) - s1)
          end
          if s1 > 0.1 * s2
             lfbaccdif1 = s1 + lfbaccdif1(a, b + 1.0)
             Exit Function
          end
       end
    end
    Call initCoeffs2
    if b + a > 2
       s1 = lfbaccdif1(b + a - 1.75, 1.75)
       a = 1.75 - b
    else
       s1 = 0.0
    end
    y1 = b - 1.0
    x1 = y1 + a
    i = UBound(coeffs2) + 3
    scale2 = 2.0 ^ -i
    scale3 = 3.0 ^ -i
    #y2 = ((y1 + 2.5) * logcf(-y1 / 2.0, i, 1.0) - (2.0 / (i - 1.0))) * scale2 + (scale3 * logcf(-y1 / 3.0, i, 1.0) + scale2 * scale2 * logcf(-y1 / 4.0, i, 1.0))
    #x2 = ((x1 + 2.5) * logcf(-x1 / 2.0, i, 1.0) - (2.0 / (i - 1.0))) * scale2 + (scale3 * logcf(-x1 / 3.0, i, 1.0) + scale2 * scale2 * logcf(-x1 / 4.0, i, 1.0))
    y2 = ((y1 + 2.5) * logcf(-y1 / 2.0, i, 1.0) - (2.0 / (i - 1.0))) * scale2 + scale3 * logcf(-y1 / 3.0, i, 1.0)
    x2 = ((x1 + 2.5) * logcf(-x1 / 2.0, i, 1.0) - (2.0 / (i - 1.0))) * scale2 + scale3 * logcf(-x1 / 3.0, i, 1.0)
    if a > 0.000006
       acc = y2 - x2  #This calculation is not accurate enough for b < 0 and a small - hence if b < 0 code above and derivative code below for small a
    else
       y3 = -(y1 + a / 2.0) / 2.0
       x3 = -(y1 + a / 2.0) / 3.0
       acc = -a * (scale2 * (logcf(y3, i, 1.0) + (y3 - 1.25) * (1.0 / (1.0 - y3) - i * logcf(y3, i + 1.0, 1.0))) - scale3 / 3.0 * ((1.0 / (1.0 - x3) - i * logcf(x3, i + 1.0, 1.0))))
    end
    For i = UBound(coeffs2) To 0 Step -1
       acc = (a * y2 - x1 * acc)
       y2 = (coeffs2(i) - y1 * y2)
       x2 = (coeffs2(i) - x1 * x2)
    Next i
    lfbaccdif1 = s1 + (y1 * y1 * acc - a * (x2 * (x1 + y1) + Onep25Minusln2Minuseulers_const))
  else
    lfbaccdif1 = [#VALUE!]
  end
end

function  logdif(pr::Float64, prob::Float64)::Float64
   Dim temp::Float64
   temp = (pr - prob) / prob
   if abs(temp) >= 0.5
      logdif = log(pr / prob)
   else
      logdif = log0(temp)
   end
end

function  cnormal(x::Float64)::Float64
#//Probability that a normal variate <= x
  Dim acc::Float64, x2::Float64, d::Float64, term::Float64, a1::Float64, a2::Float64, b1::Float64, b2::Float64, c1::Float64, c2::Float64, c3::Float64

  if (abs(x) < 1.5)
     acc = 0.0
     x2 = x * x
     term = 1.0
     d = 3.0

     while (term > sumAcc * acc)

        d = d + 2.0
        term = term * x2 / d
        acc = acc + term

     end

     acc = 1.0 + x2 / 3.0 * (1.0 + acc)
     cnormal = 0.5 + exp(-x * x * 0.5) * x * acc * OneOverSqrTwoPi
  elseif (abs(x) > 40.0)
     if (x > 0.0)
        cnormal = 1.0
     else
        cnormal = 0.0
     end
  else
     x2 = x * x
     a1 = 2.0
     b1 = x2 + 5.0
     c2 = x2 + 9.0
     a2 = a1 * c2
     b2 = b1 * c2 - 12.0
     c1 = 5.0
     c2 = c2 + 4.0

     while ((abs(a2 * b1 - a1 * b2) > abs(cfVSmall * b1 * a2)))

       c3 = c1 * (c1 + 1.0)
       a1 = c2 * a2 - c3 * a1
       b1 = c2 * b2 - c3 * b1
       c1 = c1 + 2.0
       c2 = c2 + 4.0
       c3 = c1 * (c1 + 1.0)
       a2 = c2 * a1 - c3 * a2
       b2 = c2 * b1 - c3 * b2
       c1 = c1 + 2.0
       c2 = c2 + 4.0
       if (b2 > scalefactor)
         a1 = a1 * scalefactor2
         b1 = b1 * scalefactor2
         a2 = a2 * scalefactor2
         b2 = b2 * scalefactor2
       end

     end

     if (x > 0.0)
        cnormal = 1.0 - exp(-x * x * 0.5) * OneOverSqrTwoPi * x / (x2 + 1.0 - a2 / b2)
     else
        cnormal = -exp(-x * x * 0.5) * OneOverSqrTwoPi * x / (x2 + 1.0 - a2 / b2)
     end

  end
end

function  invcnormal(p::Float64)::Float64
#//Inverse of cnormal from AS241.
#//Require p to be strictly in the range 0..1

   Dim PPND16::Float64, q::Float64, r::Float64
   q = p - 0.5
   if (abs(q) <= 0.425)
      r = 0.180625 - q * q
      PPND16 = q * (((((((a7 * r + a6) * r + a5) * r + a4) * r + a3) * r + a2) * r + a1) * r + a0) / (((((((b7 * r + b6) * r + b5) * r + b4) * r + b3) * r + b2) * r + b1) * r + 1.0)
   else
      if (q < 0.0)
         r = p
      else
         r = 1.0 - p
      end
      r = abs2(-log(r))
      if (r <= 5.0)
        r = r - 1.6
        PPND16 = (((((((c7 * r + c6) * r + c5) * r + c4) * r + c3) * r + c2) * r + c1) * r + c0) / (((((((d7 * r + d6) * r + d5) * r + d4) * r + d3) * r + d2) * r + d1) * r + 1.0)
      else
        r = r - 5.0
        PPND16 = (((((((e7 * r + e6) * r + e5) * r + e4) * r + e3) * r + e2) * r + e1) * r + e0) / (((((((f7 * r + f6) * r + f5) * r + f4) * r + f3) * r + f2) * r + f1) * r + 1.0)
      end
      if (q < 0.0)
         PPND16 = -PPND16
      end
   end
   invcnormal = PPND16
end

function pdf_lognormal(x::Float64, mean::Float64, sd::Float64)::Float64
   if (sd <= 0.0)
      pdf_lognormal = [#VALUE!]
   else
      pdf_lognormal = exp(-0.5 * ((log(x) - mean) / sd) ^ 2) / x / sd * OneOverSqrTwoPi
   end
end

function cdf_lognormal(x::Float64, mean::Float64, sd::Float64)::Float64
   if (sd <= 0.0)
      cdf_lognormal = [#VALUE!]
   else
      cdf_lognormal = cnormal((log(x) - mean) / sd)
   end
end

function comp_cdf_lognormal(x::Float64, mean::Float64, sd::Float64)::Float64
   if (sd <= 0.0)
      comp_cdf_lognormal = [#VALUE!]
   else
      comp_cdf_lognormal = cnormal(-(log(x) - mean) / sd)
   end
end

function inv_lognormal(prob::Float64, mean::Float64, sd::Float64)::Float64
   if (prob <= 0.0 || prob >= 1.0 || sd <= 0.0)
      inv_lognormal = [#VALUE!]
   else
      inv_lognormal = exp(mean + sd * invcnormal(prob))
   end
end

function comp_inv_lognormal(prob::Float64, mean::Float64, sd::Float64)::Float64
   if (prob <= 0.0 || prob >= 1.0 || sd <= 0.0)
      comp_inv_lognormal = [#VALUE!]
   else
      comp_inv_lognormal = exp(mean - sd * invcnormal(prob))
   end
end

function  tdistexp(p::Float64, q::Float64, logqk2::Float64, k::Float64, ByRef tdistDensity::Float64)::Float64
#//Special transformation of t-distribution useful for BinApprox.
#//Note approxtdistDens only used by binApprox if k > 100 or so.
   Dim sum::Float64, aki::Float64, ai::Float64, term::Float64, q1::Float64, q8::Float64
   Dim c1::Float64, c2::Float64, a1::Float64, a2::Float64, b1::Float64, b2::Float64, cadd::Float64
   Dim result::Float64, approxtdistDens::Float64

   approxtdistDens = exp(logqk2 + logfbit(k - 1.0) - 2.0 * logfbit(k * 0.5 - 1.0)) * OneOverSqrTwoPi

   if (k * p < 4.0 * q)
     sum = 0.0
     aki = k + 1.0
     ai = 3.0
     term = 1.0

     while (term > sumAcc * sum)

        ai = ai + 2.0
        aki = aki + 2.0
        term = term * aki * p / ai
        sum = sum + term

     end

     sum = 1.0 + (k + 1.0) * p * (1.0 + sum) / 3.0
     result = 0.5 - approxtdistDens * sum * abs2(k * p)
   elseif approxtdistDens = 0.0
     result = 0.0
   else
     q1 = 2.0 * (1.0 + q)
     q8 = 8.0 * q
     a1 = 1.0
     b1 = (k - 3.0) * p + 7.0
     c1 = -20.0 * q
     a2 = (k - 5.0) * p + 11.0
     b2 = a2 * b1 + c1
     cadd = -30.0 * q
     c1 = -42.0 * q
     c2 = (k - 7.0) * p + 15.0

     while ((abs(a2 * b1 - a1 * b2) > abs(cfVSmall * b1 * a2)))

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
       if (abs(b2) > scalefactor)
         a1 = a1 * scalefactor2
         b1 = b1 * scalefactor2
         a2 = a2 * scalefactor2
         b2 = b2 * scalefactor2
       elseif (abs(b2) < scalefactor2)
         a1 = a1 * scalefactor
         b1 = b1 * scalefactor
         a2 = a2 * scalefactor
         b2 = b2 * scalefactor
       end
     end

     result = approxtdistDens * (1.0 - q / ((k - 1.0) * p + 3.0 - 6.0 * q * a2 / b2)) / abs2(k * p)
   end
   tdistDensity = approxtdistDens * abs2(q)
   tdistexp = result
end

function  tdist(x::Float64, k::Float64, tdistDensity::Float64)::Float64
#//Probability that variate from t-distribution with k degress of freedom <= x
   Dim x2::Float64, k2::Float64, logterm::Float64, a::Float64, r::Float64, c5::Float64

   if abs(x) >= min(1.0, k)
      k2 = k / x
      x2 = x + k2
      k2 = k2 / x2
      x2 = x / x2
   else
      x2 = x * x
      k2 = k + x2
      x2 = x2 / k2
      k2 = k / k2
   end
   if (k > 1E+30)
      tdist = cnormal(x)
      tdistDensity = exp(-x * x * 0.5) * OneOverSqrTwoPi
   else
      a = k * 0.5
      if (k2 < cSmall)
        logterm = (log(k) - 2.0 * log(abs(x)))
      elseif (abs(x2) < 0.5)
        logterm = log0(-x2)
      else
        logterm = log(k2)
      end
      if (k >= 1.0)
         logterm = logterm * a
         if (x < 0.0)
           tdist = tdistexp(x2, k2, logterm, k, tdistDensity)
         else
           tdist = 1.0 - tdistexp(x2, k2, logterm, k, tdistDensity)
         end
         Exit Function
      end
      c5 = -1.0 / (k + 2.0)
      tdistDensity = exp((a + 0.5) * logterm + a * log1(c5) - c5 + lfbaccdif1(0.5, a - 0.5)) * abs2(a / ((1.0 + a))) * OneOverSqrTwoPi
      if (k2 < cSmall)
        r = (a + 1.0) * log1(a / 1.5) - lfbaccdif1(a, 0.5) - lngammaexpansion(a)
        r = r + a * ((a - 0.5) / 1.5 + Log1p5 + (log(k) - 2.0 * log(abs(x))))
        r = exp(r) * (0.25 / (a + 0.5))
        if x < 0.0
           tdist = r
        else
           tdist = 1.0 - r
        end
      elseif (x < 0.0)
        if x2 < k2
          tdist = 0.5 * compbeta(x2, 0.5, a)
        else
          tdist = 0.5 * beta(k2, a, 0.5)
        end
      else
        if x2 < k2
          tdist = 0.5 * (1.0 + beta(x2, 0.5, a))
        else
          tdist = 0.5 * (1.0 + compbeta(k2, a, 0.5))
        end
      end
   end
end

function  BetterThanTailApprox(prob::Float64, df::Float64)::Bool
if df <= 2
   BetterThanTailApprox = prob > 0.25 * exp((1.0 - df) * 1.78514841051368)
elseif df <= 5
   BetterThanTailApprox = prob > 0.045 * exp((2.0 - df) * 1.30400766847605)
elseif df <= 20
   BetterThanTailApprox = prob > 0.0009 * exp((5.0 - df) * 0.921034037197618)
else
   BetterThanTailApprox = prob > 0.0000000009 * exp((20.0 - df) * 0.690775527898214)
end
end

function  invtdist(prob::Float64, df::Float64)::Float64
#//Inverse of tdist
#//Require prob to be in the range 0..1 df should be positive
  Dim xn::Float64, xn2::Float64, tp::Float64, tpDif::Float64, tprob::Float64, a::Float64, pr::Float64, lpr::Float64, small::Float64, smalllpr::Float64, tdistDensity::Float64
  if prob > 0.5
     pr = 1.0 - prob
  else
     pr = prob
  end
  lpr = -log(pr)
  small = 0.00000000000001
  smalllpr = small * lpr * pr
  if pr >= 0.5 || df >= 1.0 && BetterThanTailApprox(pr, df)
#// Will divide by 0 if tp so small that tdistDensity underflows. !a problem if prob > cSmall
     xn = invcnormal(pr)
     xn2 = xn * xn
#//Initial approximation is given in http://digital.library.adelaide.edu.au/coll/special//fisher/281.pdf. The modified NR correction then gets it right.
     tp = (((((27.0 * xn2 + 339.0) * xn2 + 930.0) * xn2 - 1782.0) * xn2 - 765.0) * xn2 + 17955.0) / (368640.0 * df)
     tp = (tp + ((((79.0 * xn2 + 776.0) * xn2 + 1482.0) * xn2 - 1920.0) * xn2 - 945.0) / 92160.0) / df
     tp = (tp + (((3.0 * xn2 + 19.0) * xn2 + 17.0) * xn2 - 15.0) / 384.0) / df
     tp = (tp + ((5.0 * xn2 + 16) * xn2 + 3.0) / 96.0) / df
     tp = (tp + (xn2 + 1.0) / 4.0) / df
     tp = xn * (1.0 + tp)
     tprob = 0.0
     tpDif = 1.0 + abs(tp)
  elseif df < 1.0
     a = df / 2.0
     tp = (a + 1.0) * log1(a / 1.5) - lfbaccdif1(a, 0.5) - lngammaexpansion(a)
     tp = ((a - 0.5) / 1.5 + Log1p5 + log(df)) / 2.0 + (tp - log(4.0 * pr * (a + 0.5))) / df
     tp = -exp(tp)
     tprob = tdist(tp, df, tdistDensity)
     if tdistDensity < nearly_zero
        tpDif = 0.0
     else
        tpDif = (tprob / tdistDensity) * log0((tprob - pr) / pr)
        tp = tp - tpDif
     end
  else
     tp = tdist(0, df, tdistDensity) #Marginally quicker to get tdistDensity for integral df
     tp = exp(-log(abs2(df) * pr / tdistDensity) / df)
     if df >= 2
        tp = -abs2(df * (tp * tp - 1.0))
     else
        tp = -abs2(df) * abs2(tp - 1.0) * abs2(tp + 1.0)
     end
     tpDif = tp / df
     tpDif = -log0((0.5 - 1.0 / (df + 2)) / (1.0 + tpDif * tp)) * (tpDif + 1.0 / tp)
     tp = tp - tpDif
     tprob = tdist(tp, df, tdistDensity)
     if tdistDensity < nearly_zero
        tpDif = 0.0
     else
        tpDif = (tprob / tdistDensity) * log0((tprob - pr) / pr)
        tp = tp - tpDif
     end
  end
  while (abs(tprob - pr) > smalllpr && abs(tpDif) > small * (1.0 + abs(tp)))
     tprob = tdist(tp, df, tdistDensity)
     tpDif = (tprob / tdistDensity) * log0((tprob - pr) / pr)
     tp = tp - tpDif
  end
  invtdist = tp
  if prob > 0.5 invtdist = -invtdist
end

function  poissonTerm(i::Float64, n::Float64, diffFromMean::Float64, logAdd::Float64)::Float64
#//Probability that poisson variate with mean n has value i (diffFromMean = n-i)
   Dim c2::Float64, c3::Float64
   Dim logpoissonTerm::Float64, c1::Float64

   if ((i <= -1.0) || (n < 0.0))
      if (i = 0.0)
         poissonTerm = exp(logAdd)
      else
         poissonTerm = 0.0
      end
   elseif ((i < 0.0) && (n = 0.0))
      poissonTerm = [#VALUE!]
   else
     c3 = i
     c2 = c3 + 1.0
     c1 = (diffFromMean - 1.0) / c2

     if (c1 < minLog1Value)
        if (i = 0.0)
          logpoissonTerm = -n
          poissonTerm = exp(logpoissonTerm + logAdd)
        else
          On Error GoTo ptiszero
          logpoissonTerm = (c3 * log(n / c2) - (diffFromMean - 1.0)) - logfbit(c3)
          poissonTerm = exp(logpoissonTerm + logAdd) / abs2(c2) * OneOverSqrTwoPi
          Exit Function
ptiszero: poissonTerm = 0.0
          Exit Function
        end
     else
       logpoissonTerm = c3 * log1(c1) - c1 - logfbit(c3)
       poissonTerm = exp(logpoissonTerm + logAdd) / abs2(c2) * OneOverSqrTwoPi
     end
   end
end

function  poisson1(i::Float64, n::Float64, diffFromMean::Float64)::Float64
#//Probability that poisson variate with mean n has value <= i (diffFromMean = n-i)
#//For negative values of i (used for calculating the cumlative gamma distribution) there#s a really nasty interpretation!
#//1-gamma(n,i) is calculated as poisson1(-i,n,0) since we need an accurate version of i rather than i-1.
#//Uses a simplified version of Legendre#s continued fraction.
   Dim prob::Float64, exact::Bool
   if ((i >= 0.0) && (n <= 0.0))
      exact = true
      prob = 1.0
   elseif ((i > -1.0) && (n <= 0.0))
      exact = true
      prob = 0.0
   elseif ((i > -1.0) && (i < 0.0))
      i = -i
      exact = false
      prob = poissonTerm(i, n, n - i, 0.0) * i / n
      i = i - 1.0
      diffFromMean = n - i
   else
      exact = ((i <= -1.0) || (n < 0.0))
      prob = poissonTerm(i, n, diffFromMean, 0.0)
   end
   if (exact || prob = 0.0)
      poisson1 = prob
      Exit Function
   end

   Dim a1::Float64, a2::Float64, b1::Float64, b2::Float64, c1::Float64, c2::Float64, c3::Float64, c4::Float64, cfValue::Float64
   Dim njj As Long, numb As Long
   Dim sumAlways As Long, sumFactor As Long
   sumAlways = 0
   sumFactor = 6
   a1 = 0.0
   if (i > sumAlways)
      numb = Int(sumFactor * exp(log(n) / 3))
      numb = max(0, Int(numb - diffFromMean))
      if (numb > i)
         numb = Int(i)
      end
   else
      numb = max(0, Int(i))
   end

   b1 = 1.0
   a2 = i - numb
   b2 = diffFromMean + (numb + 1.0)
   c1 = 0.0
   c2 = a2
   c4 = b2
   if c2 < 0.0
      cfValue = cfVSmall
   else
      cfValue = cfSmall
   end
   while ((abs(a2 * b1 - a1 * b2) > abs(cfValue * b1 * a2)))

       c1 = c1 + 1.0
       c2 = c2 - 1.0
       c3 = c1 * c2
       c4 = c4 + 2.0
       a1 = c4 * a2 + c3 * a1
       b1 = c4 * b2 + c3 * b1
       c1 = c1 + 1.0
       c2 = c2 - 1.0
       c3 = c1 * c2
       c4 = c4 + 2.0
       a2 = c4 * a1 + c3 * a2
       b2 = c4 * b1 + c3 * b2
       if (b2 > scalefactor)
         a1 = a1 * scalefactor2
         b1 = b1 * scalefactor2
         a2 = a2 * scalefactor2
         b2 = b2 * scalefactor2
       end
       if c2 < 0.0 && cfValue > cfVSmall
          cfValue = cfVSmall
       end
   end

   a1 = a2 / b2

   c1 = i - numb + 1.0
   For njj = 1 To numb
     a1 = (1.0 + a1) * (c1 / n)
     c1 = c1 + 1.0
   Next njj

   poisson1 = (1.0 + a1) * prob
end

function  poisson2(i::Float64, n::Float64, diffFromMean::Float64)::Float64
#//Probability that poisson variate with mean n has value >= i (diffFromMean = n-i)
   Dim prob::Float64, exact::Bool
   if ((i <= 0.0) && (n <= 0.0))
      exact = true
      prob = 1.0
   else
      exact = false
      prob = poissonTerm(i, n, diffFromMean, 0.0)
   end
   if (exact || prob = 0.0)
      poisson2 = prob
      Exit Function
   end

   Dim a1::Float64, a2::Float64, b1::Float64, b2::Float64, c1::Float64, c2::Float64
   Dim njj As Long, numb As Long
   const sumFactor = 6
   numb = Int(sumFactor * exp(log(n) / 3))
   numb = max(0, Int(diffFromMean + numb))

   a1 = 0.0
   b1 = 1.0
   a2 = n
   b2 = (numb + 1.0) - diffFromMean
   c1 = 0.0
   c2 = b2

   while ((abs(a2 * b1 - a1 * b2) > abs(cfSmall * b1 * a2)))

      c1 = c1 + n
      c2 = c2 + 1.0
      a1 = c2 * a2 + c1 * a1
      b1 = c2 * b2 + c1 * b1
      c1 = c1 + n
      c2 = c2 + 1.0
      a2 = c2 * a1 + c1 * a2
      b2 = c2 * b1 + c1 * b2
      if (b2 > scalefactor)
        a1 = a1 * scalefactor2
        b1 = b1 * scalefactor2
        a2 = a2 * scalefactor2
        b2 = b2 * scalefactor2
      end
   end

   a1 = a2 / b2

   c1 = i + numb
   For njj = 1 To numb
     a1 = (1.0 + a1) * (n / c1)
     c1 = c1 - 1.0
   Next

   poisson2 = (1.0 + a1) * prob

end

function  poissonApprox(j::Float64, diffFromMean::Float64, comp::Bool)::Float64
#//Asymptotic expansion to calculate the probability that poisson variate has value <= j (diffFromMean = mean-j). if comp then calulate 1-probability.
#//cf. http://members.aol.com/iandjmsmith/PoissonApprox.htm
Dim pt::Float64, s2pt::Float64, res1::Float64, res2::Float64, elfb::Float64, term::Float64
Dim ig2::Float64, ig3::Float64, ig4::Float64, ig5::Float64, ig6::Float64, ig7::Float64, ig8::Float64
Dim ig05::Float64, ig25::Float64, ig35::Float64, ig45::Float64, ig55::Float64, ig65::Float64, ig75::Float64

pt = -log1(diffFromMean / j)
s2pt = abs2(2.0 * j * pt)

ig2 = 1.0 / j + pt
term = pt * pt * 0.5
ig3 = ig2 / j + term
term = term * pt / 3.0
ig4 = ig3 / j + term
term = term * pt / 4.0
ig5 = ig4 / j + term
term = term * pt / 5.0
ig6 = ig5 / j + term
term = term * pt / 6.0
ig7 = ig6 / j + term
term = term * pt / 7.0
ig8 = ig7 / j + term

ig05 = cnormal(-s2pt)
term = pt * twoThirds
ig25 = 1.0 / j + term
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
res1 = (((((((ig8 * coef8 + ig7 * coef7) + ig6 * coef6) + ig5 * coef5) + ig4 * coef4) + ig3 * coef3) + ig2 * coef2) + coef1) * abs2(j)
res2 = ((((((ig75 * coef75 + ig65 * coef65) + ig55 * coef55) + ig45 * coef45) + ig35 * coef35) + ig25 * coef25) + coef15) * s2pt

if (comp)
   if (diffFromMean < 0.0)
      poissonApprox = ig05 - (res1 - res2) * exp(-j * pt) * OneOverSqrTwoPi / elfb
   else
      poissonApprox = (1.0 - ig05) - (res1 + res2) * exp(-j * pt) * OneOverSqrTwoPi / elfb
   end
elseif (diffFromMean < 0.0)
   poissonApprox = (1.0 - ig05) + (res1 - res2) * exp(-j * pt) * OneOverSqrTwoPi / elfb
else
   poissonApprox = ig05 + (res1 + res2) * exp(-j * pt) * OneOverSqrTwoPi / elfb
end
end

function  cpoisson(k::Float64, lambda::Float64, dfm::Float64)::Float64
#//Probability that poisson variate with mean lambda has value <= k (diffFromMean = lambda-k) calculated by various methods.
   if ((k >= 21.0) && (abs(dfm) < (0.3 * k)))
      cpoisson = poissonApprox(k, dfm, false)
   elseif ((lambda > k) && (lambda >= 1.0))
      cpoisson = poisson1(k, lambda, dfm)
   else
      cpoisson = 1.0 - poisson2(k + 1.0, lambda, dfm - 1.0)
   end
end

function  comppoisson(k::Float64, lambda::Float64, dfm::Float64)::Float64
#//Probability that poisson variate with mean lambda has value > k (diffFromMean = lambda-k) calculated by various methods.
   if ((k >= 21.0) && (abs(dfm) < (0.3 * k)))
      comppoisson = poissonApprox(k, dfm, true)
   elseif ((lambda > k) && (lambda >= 1.0))
      comppoisson = 1.0 - poisson1(k, lambda, dfm)
   else
      comppoisson = poisson2(k + 1.0, lambda, dfm - 1.0)
   end
end

function  invpoisson(k::Float64, prob::Float64)::Float64
#//Inverse of poisson. Calculates mean such that poisson(k,mean,mean-k)=prob.
#//Require prob to be in the range 0..1, k should be -1/2 or non-negative
   if (k = 0.0)
      invpoisson = -log(prob + 9.99988867182683E-321)
   elseif (prob > 0.5)
      invpoisson = invcomppoisson(k, 1.0 - prob)
   else #/*if (k > 0.0)*/ then
      Dim temp2::Float64, xp::Float64, dfm::Float64, q::Float64, qdif::Float64, lpr::Float64, small::Float64, smalllpr::Float64
      lpr = -log(prob)
      small = 0.00000000000001
      smalllpr = small * lpr * prob
      xp = invcnormal(prob)
      dfm = xp * (0.5 * xp - abs2(k + (0.5 * xp) ^ 2))
      q = -1.0
      qdif = -dfm
      if abs(qdif) < 1.0
         qdif = 1.0
      elseif (k > 1E+50)
         invpoisson = k
         Exit Function
      end
      while ((abs(q - prob) > smalllpr) && (abs(qdif) > (1.0 + abs(dfm)) * small))
         q = cpoisson(k, k + dfm, dfm)
         if (q = 0.0)
             qdif = qdif / 2.0
             dfm = dfm + qdif
             q = -1.0
         else
            temp2 = poissonTerm(k, k + dfm, dfm, 0.0)
            if (temp2 = 0.0)
               qdif = qdif / 2.0
               dfm = dfm + qdif
               q = -1.0
            else
               qdif = -2.0 * q * logdif(q, prob) / (1.0 + abs2(log(prob) / log(q))) / temp2
               if (qdif > k + dfm)
                  qdif = dfm / 2.0
                  dfm = dfm - qdif
                  q = -1.0
               else
                  dfm = dfm - qdif
               end
            end
         end
      end
      invpoisson = k + dfm
   end
end

function  invcomppoisson(k::Float64, prob::Float64)::Float64
#//Inverse of comppoisson. Calculates mean such that comppoisson(k,mean,mean-k)=prob.
#//Require prob to be in the range 0..1, k should be -1/2 or non-negative
   if (prob > 0.5)
      invcomppoisson = invpoisson(k, 1.0 - prob)
   elseif (k = 0.0)
      invcomppoisson = -log0(-prob)
   else #/*if (k > 0.0)*/ then
      Dim temp2::Float64, xp::Float64, dfm::Float64, q::Float64, qdif::Float64, lambda::Float64, qdifset::Bool, lpr::Float64, small::Float64, smalllpr::Float64
      lpr = -log(prob)
      small = 0.00000000000001
      smalllpr = small * lpr * prob
      xp = invcnormal(prob)
      dfm = xp * (0.5 * xp + abs2(k + (0.5 * xp) ^ 2))
      lambda = k + dfm
      if ((lambda < 1.0) && (k < 40.0))
         lambda = exp(log(prob / poissonTerm(k + 1.0, 1.0, -k, 0.0)) / (k + 1.0))
         dfm = lambda - k
      elseif (k > 1E+50)
         invcomppoisson = lambda
         Exit Function
      end
      q = -1.0
      qdif = lambda
      qdifset = false
      while ((abs(q - prob) > smalllpr) && (abs(qdif) > min(lambda, abs(dfm)) * small))
         q = comppoisson(k, lambda, dfm)
         if (q = 0.0)
            if qdifset
               qdif = qdif / 2.0
               dfm = dfm + qdif
               lambda = lambda + qdif
            else
               lambda = 2.0 * lambda
               qdif = qdif * 2.0
               dfm = lambda - k
            end
            q = -1.0
         else
            temp2 = poissonTerm(k, lambda, dfm, 0.0)
            if (temp2 = 0.0)
               if qdifset
                  qdif = qdif / 2.0
                  dfm = dfm + qdif
                  lambda = lambda + qdif
               else
                  lambda = 2.0 * lambda
                  qdif = qdif * 2.0
                  dfm = lambda - k
               end
               q = -1.0
            else
               qdif = 2.0 * q * logdif(q, prob) / (1.0 + abs2(log(prob) / log(q))) / temp2
               if (qdif > lambda)
                  lambda = lambda / 10.0
                  qdif = dfm
                  dfm = lambda - k
                  qdif = qdif - dfm
                  q = -1.0
               else
                  lambda = lambda - qdif
                  dfm = dfm - qdif
               end
               qdifset = true
            end
         end
         if (abs(dfm) > lambda)
            dfm = lambda - k
         else
            lambda = k + dfm
         end
      end
      invcomppoisson = lambda
   end
end

function  binomialTerm(i::Float64, j::Float64, p::Float64, q::Float64, diffFromMean::Float64, logAdd::Float64)::Float64
#//Probability that binomial variate with sample size i+j and event prob p (=1-q) has value i (diffFromMean = (i+j)*p-i)
   Dim c1::Float64, c2::Float64, c3::Float64
   Dim c4::Float64, c5::Float64, c6::Float64, ps::Float64, logbinomialTerm::Float64, dfm::Float64
   if ((i = 0.0) && (j <= 0.0))
      binomialTerm = exp(logAdd)
   elseif ((i <= -1.0) || (j < 0.0))
      binomialTerm = 0.0
   else
      if (p < q)
         c2 = i
         c3 = j
         ps = p
         dfm = diffFromMean
      else
         c3 = i
         c2 = j
         ps = q
         dfm = -diffFromMean
      end

      c5 = (dfm - (1.0 - ps)) / (c2 + 1.0)
      c6 = -(dfm + ps) / (c3 + 1.0)

      if (c5 < minLog1Value)
         if (c2 = 0.0)
            logbinomialTerm = c3 * log0(-ps)
            binomialTerm = exp(logbinomialTerm + logAdd)
         elseif ((ps = 0.0) && (c2 > 0.0))
            binomialTerm = 0.0
         else
            c1 = (i + 1.0) + j
            #c4 = logfbit(i + j) - logfbit(i) - logfbit(j)
            #logbinomialTerm = c4 + c2 * (log((ps * c1) / (c2 + 1.0)) - c5) - c5 + c3 * log1(c6) - c6
            c4 = lfbaccdif1(j, i) + logfbit(j)
            logbinomialTerm = c2 * (log((ps * c1) / (c2 + 1.0)) - c5) - c5 + c3 * log1(c6) - c6 - c4
            binomialTerm = exp(logbinomialTerm + logAdd) * abs2(c1 / ((c2 + 1.0) * (c3 + 1.0))) * OneOverSqrTwoPi
         end
      else
         #c4 = logfbit(i + j) - logfbit(i) - logfbit(j)
         #logbinomialTerm = c4 + (c2 * log1(c5) - c5) + (c3 * log1(c6) - c6)
         c4 = lfbaccdif1(j, i) + logfbit(j)
         logbinomialTerm = (c2 * log1(c5) - c5) + (c3 * log1(c6) - c6) - c4
         binomialTerm = exp(logbinomialTerm + logAdd) * abs2((1.0 + j / (i + 1.0)) / (j + 1.0)) * OneOverSqrTwoPi
      end
   end
end

function  binomialcf(ii::Float64, jj::Float64, pp::Float64, qq::Float64, diffFromMean::Float64, comp::Bool)::Float64
#//Probability that binomial variate with sample size ii+jj and event prob pp (=1-qq) has value <=i (diffFromMean = (ii+jj)*pp-ii). if comp the returns 1 - probability.
Dim prob::Float64, p::Float64, q::Float64, a1::Float64, a2::Float64, b1::Float64, b2::Float64
Dim c1::Float64, c2::Float64, c3::Float64, c4::Float64, n1::Float64, q1::Float64, dfm::Float64
Dim i::Float64, j::Float64, ni::Float64, nj::Float64, numb::Float64, ip1::Float64, cfValue::Float64
Dim swapped::Bool, exact::Bool

  if ((ii > -1.0) && (ii < 0.0))
     ip1 = -ii
     ii = ip1 - 1.0
  else
     ip1 = ii + 1.0
  end
  n1 = (ii + 3.0) + jj
  if ii < 0.0
     cfValue = cfVSmall
     swapped = false
  elseif pp > qq
     cfValue = cfSmall
     swapped = n1 * qq >= jj + 1.0
  else
     cfValue = cfSmall
     swapped = n1 * pp <= ii + 2.0
  end
  if Not swapped
    i = ii
    j = jj
    p = pp
    q = qq
    dfm = diffFromMean
  else
    j = ip1
    ip1 = jj
    i = jj - 1.0
    p = qq
    q = pp
    dfm = 1.0 - diffFromMean
  end
  if ((i > -1.0) && ((j <= 0.0) || (p = 0.0)))
     exact = true
     prob = 1.0
  elseif ((i > -1.0) && (i < 0.0) || (i = -1.0) && (ip1 > 0.0))
     exact = false
     prob = binomialTerm(ip1, j, p, q, (ip1 + j) * p - ip1, 0.0) * ip1 / ((ip1 + j) * p)
     dfm = (i + j) * p - i
  else
     exact = ((i = 0.0) && (j <= 0.0)) || ((i <= -1.0) || (j < 0.0))
     prob = binomialTerm(i, j, p, q, dfm, 0.0)
  end
  if (exact) || (prob = 0.0)
     if (swapped = comp)
        binomialcf = prob
     else
        binomialcf = 1.0 - prob
     end
     Exit Function
  end

  Dim sumAlways As Long, sumFactor As Long
  sumAlways = 0
  sumFactor = 6
  a1 = 0.0
  if (i > sumAlways)
     numb = Int(sumFactor * abs2(p + 0.5) * exp(log(n1 * p * q) / 3))
     numb = Int(numb - dfm)
     if (numb > i)
        numb = Int(i)
     end
  else
     numb = Int(i)
  end
  if (numb < 0.0)
     numb = 0.0
  end

  b1 = 1.0
  q1 = q + 1.0
  a2 = (i - numb) * q
  b2 = dfm + numb + 1.0
  c1 = 0.0
  c2 = a2
  c4 = b2
  while ((abs(a2 * b1 - a1 * b2) > abs(cfValue * b1 * a2)))

    c1 = c1 + 1.0
    c2 = c2 - q
    c3 = c1 * c2
    c4 = c4 + q1
    a1 = c4 * a2 + c3 * a1
    b1 = c4 * b2 + c3 * b1
    c1 = c1 + 1.0
    c2 = c2 - q
    c3 = c1 * c2
    c4 = c4 + q1
    a2 = c4 * a1 + c3 * a2
    b2 = c4 * b1 + c3 * b2
    if (abs(b2) > scalefactor)
      a1 = a1 * scalefactor2
      b1 = b1 * scalefactor2
      a2 = a2 * scalefactor2
      b2 = b2 * scalefactor2
    elseif (abs(b2) < scalefactor2)
      a1 = a1 * scalefactor
      b1 = b1 * scalefactor
      a2 = a2 * scalefactor
      b2 = b2 * scalefactor
    end
    if c2 < 0.0 && cfValue > cfVSmall
       cfValue = cfVSmall
    end
  end
  a1 = a2 / b2

  ni = (i - numb + 1.0) * q
  nj = (j + numb) * p
  while (numb > 0.0)
     a1 = (1.0 + a1) * (ni / nj)
     ni = ni + q
     nj = nj - p
     numb = numb - 1.0
  end

  a1 = (1.0 + a1) * prob
  if (swapped = comp)
     binomialcf = a1
  else
     binomialcf = 1.0 - a1
  end

end

function  binApprox(a::Float64, b::Float64, diffFromMean::Float64, comp::Bool)::Float64
#//Asymptotic expansion to calculate the probability that binomial variate has value <= a (diffFromMean = (a+b)*p-a). if comp then calulate 1-probability.
#//cf. http://members.aol.com/iandjmsmith/BinomialApprox.htm
Dim n::Float64, n1::Float64
Dim pq1::Float64, mfac::Float64, res::Float64, tp::Float64, lval::Float64, lvv::Float64, temp::Float64
Dim ib05::Float64, ib15::Float64, ib25::Float64, ib35::Float64, ib45::Float64, ib55::Float64, ib65::Float64
Dim ib2::Float64, ib3::Float64, ib4::Float64, ib5::Float64, ib6::Float64, ib7::Float64
Dim elfb::Float64, coef15::Float64, coef25::Float64, coef35::Float64, coef45::Float64, coef55::Float64, coef65::Float64
Dim coef2::Float64, coef3::Float64, coef4::Float64, coef5::Float64, coef6::Float64, coef7::Float64
Dim tdistDensity::Float64, approxtdistDens::Float64

n = a + b
n1 = n + 1.0
lvv = (b + diffFromMean) / n1 - diffFromMean
lval = (a * log1(lvv / a) + b * log1(-lvv / b)) / n
tp = -expm1(lval)

pq1 = (a / n) * (b / n)

coef15 = (-17.0 * pq1 + 2.0) / 24.0
coef25 = ((-503.0 * pq1 + 76.0) * pq1 + 4.0) / 1152.0
coef35 = (((-315733.0 * pq1 + 53310.0) * pq1 + 8196.0) * pq1 - 1112.0) / 414720.0
coef45 = (4059192.0 + pq1 * (15386296.0 - 85262251.0 * pq1))
coef45 = (-9136.0 + pq1 * (-697376 + pq1 * coef45)) / 39813120.0
coef55 = (3904584040.0 + pq1 * (10438368262.0 - 55253161559.0 * pq1))
coef55 = (5244128.0 + pq1 * (-43679536.0 + pq1 * (-703410640.0 + pq1 * coef55))) / 6688604160.0
coef65 = (-3242780782432.0 + pq1 * (18320560326516.0 + pq1 * (38020748623980.0 - 194479285104469.0 * pq1)))
coef65 = (335796416.0 + pq1 * (61701376704.0 + pq1 * (-433635420336.0 + pq1 * coef65))) / 4815794995200.0
elfb = (((((coef65 / ((n + 6.5) * pq1) + coef55) / ((n + 5.5) * pq1) + coef45) / ((n + 4.5) * pq1) + coef35) / ((n + 3.5) * pq1) + coef25) / ((n + 2.5) * pq1) + coef15) / ((n + 1.5) * pq1) + 1.0

coef2 = (-pq1 - 2.0) / 135.0
coef3 = ((-44.0 * pq1 - 86.0) * pq1 + 4.0) / 2835.0
coef4 = (((-404.0 * pq1 - 786.0) * pq1 + 48.0) * pq1 + 8.0) / 8505.0
coef5 = (((((-2421272.0 * pq1 - 4721524.0) * pq1 + 302244.0) * pq1) + 118160.0) * pq1 - 4496.0) / 12629925.0
coef6 = ((((((-473759128.0 * pq1 - 928767700.0) * pq1 + 57300188.0) * pq1) + 38704888.0) * pq1 - 1870064.0) * pq1 - 167072.0) / 492567075.0
coef7 = (((((((-8530742848.0 * pq1 - 16836643200.0) * pq1 + 954602040.0) * pq1) + 990295352.0) * pq1 - 44963088.0) * pq1 - 11596512.0) * pq1 + 349376.0) / 1477701225.0

ib05 = tdistexp(tp, 1.0 - tp, n1 * lval, 2.0 * n1, tdistDensity)
mfac = n1 * tp
ib15 = abs2(2.0 * mfac)

if (mfac > 1E+50)
   ib2 = (1.0 + mfac) / (n + 2.0)
   mfac = mfac * tp / 2.0
   ib3 = (ib2 + mfac) / (n + 3.0)
   mfac = mfac * tp / 3.0
   ib4 = (ib3 + mfac) / (n + 4.0)
   mfac = mfac * tp / 4.0
   ib5 = (ib4 + mfac) / (n + 5.0)
   mfac = mfac * tp / 5.0
   ib6 = (ib5 + mfac) / (n + 6.0)
   mfac = mfac * tp / 6.0
   ib7 = (ib6 + mfac) / (n + 7.0)
   res = (ib2 * coef2 + (ib3 * coef3 + (ib4 * coef4 + (ib5 * coef5 + (ib6 * coef6 + ib7 * coef7 / pq1) / pq1) / pq1) / pq1) / pq1) / pq1

   mfac = (n + 1.5) * tp * twoThirds
   ib25 = (1.0 + mfac) / (n + 2.5)
   mfac = mfac * tp * twoFifths
   ib35 = (ib25 + mfac) / (n + 3.5)
   mfac = mfac * tp * twoSevenths
   ib45 = (ib35 + mfac) / (n + 4.5)
   mfac = mfac * tp * twoNinths
   ib55 = (ib45 + mfac) / (n + 5.5)
   mfac = mfac * tp * twoElevenths
   ib65 = (ib55 + mfac) / (n + 6.5)
   temp = (((((coef65 * ib65 / pq1 + coef55 * ib55) / pq1 + coef45 * ib45) / pq1 + coef35 * ib35) / pq1 + coef25 * ib25) / pq1 + coef15)
else
   ib2 = 1.0 + mfac
   mfac = mfac * (n + 2.0) * tp / 2.0
   ib3 = ib2 + mfac
   mfac = mfac * (n + 3.0) * tp / 3.0
   ib4 = ib3 + mfac
   mfac = mfac * (n + 4.0) * tp / 4.0
   ib5 = ib4 + mfac
   mfac = mfac * (n + 5.0) * tp / 5.0
   ib6 = ib5 + mfac
   mfac = mfac * (n + 6.0) * tp / 6.0
   ib7 = ib6 + mfac
   res = (ib2 * coef2 + (ib3 * coef3 + (ib4 * coef4 + (ib5 * coef5 + (ib6 * coef6 + ib7 * coef7 / ((n + 7.0) * pq1)) / ((n + 6.0) * pq1)) / ((n + 5.0) * pq1)) / ((n + 4.0) * pq1)) / ((n + 3.0) * pq1)) / ((n + 2.0) * pq1)

   mfac = (n + 1.5) * tp * twoThirds
   ib25 = 1.0 + mfac
   mfac = mfac * (n + 2.5) * tp * twoFifths
   ib35 = ib25 + mfac
   mfac = mfac * (n + 3.5) * tp * twoSevenths
   ib45 = ib35 + mfac
   mfac = mfac * (n + 4.5) * tp * twoNinths
   ib55 = ib45 + mfac
   mfac = mfac * (n + 5.5) * tp * twoElevenths
   ib65 = ib55 + mfac
   temp = (((((coef65 * ib65 / ((n + 6.5) * pq1) + coef55 * ib55) / ((n + 5.5) * pq1) + coef45 * ib45) / ((n + 4.5) * pq1) + coef35 * ib35) / ((n + 3.5) * pq1) + coef25 * ib25) / ((n + 2.5) * pq1) + coef15)
end

approxtdistDens = tdistDensity / abs2(1.0 - tp)
temp = ib15 * temp / ((n + 1.5) * pq1)
res = (oneThird + res) * 2.0 * (a - b) / (n * abs2(n1 * pq1))
if (comp)
   if (lvv > 0.0)
      binApprox = ib05 - (res - temp) * approxtdistDens / elfb
   else
      binApprox = (1.0 - ib05) - (res + temp) * approxtdistDens / elfb
   end
elseif (lvv > 0.0)
   binApprox = (1.0 - ib05) + (res - temp) * approxtdistDens / elfb
else
   binApprox = ib05 + (res + temp) * approxtdistDens / elfb
end
end

function  binomial(ii::Float64, jj::Float64, pp::Float64, qq::Float64, diffFromMean::Float64)::Float64
#//Probability that binomial variate with sample size ii+jj and event prob pp (=1-qq) has value <=i (diffFromMean = (ii+jj)*pp-ii).
   Dim mij::Float64
   mij = min(ii, jj)
   if ((mij > 50.0) && (abs(diffFromMean) < (0.1 * mij)))
      binomial = binApprox(jj - 1.0, ii, diffFromMean, false)
   else
      binomial = binomialcf(ii, jj, pp, qq, diffFromMean, false)
   end
end

function  compbinomial(ii::Float64, jj::Float64, pp::Float64, qq::Float64, diffFromMean::Float64)::Float64
#//Probability that binomial variate with sample size ii+jj and event prob pp (=1-qq) has value >i (diffFromMean = (ii+jj)*pp-ii).
   Dim mij::Float64
   mij = min(ii, jj)
   if ((mij > 50.0) && (abs(diffFromMean) < (0.1 * mij)))
       compbinomial = binApprox(jj - 1.0, ii, diffFromMean, true)
   else
       compbinomial = binomialcf(ii, jj, pp, qq, diffFromMean, true)
   end
end

function  invbinom(k::Float64, m::Float64, prob::Float64, ByRef oneMinusP::Float64)::Float64
#//Inverse of binomial. Delivers event probability p (q held in oneMinusP in case required) so that binomial(k,m,p,oneMinusp,dfm) = prob.
#//Note that dfm is calculated accurately but never made available outside of this routine.
#//Require prob to be in the range 0..1, m should be positive and k should be >= 0
   Dim temp1::Float64, temp2::Float64
   if (prob > 0.5)
      temp2 = invcompbinom(k, m, 1.0 - prob, oneMinusP)
   else
      temp1 = invcompbinom(m - 1.0, k + 1.0, prob, oneMinusP)
      temp2 = oneMinusP
      oneMinusP = temp1
   end
   invbinom = temp2
end

function  invcompbinom(k::Float64, m::Float64, prob::Float64, ByRef oneMinusP::Float64)::Float64
#//Inverse of compbinomial. Delivers event probability p (q held in oneMinusP in case required) so that compbinomial(k,m,p,oneMinusp,dfm) = prob.
#//Note that dfm is calculated accurately but never made available outside of this routine.
#//Require prob to be in the range 0..1, m should be positive and k should be >= -0.5
Dim xp::Float64, xp2::Float64, dfm::Float64, n::Float64, p::Float64, q::Float64, pr::Float64, dif::Float64, temp::Float64, temp2::Float64, result::Float64, lpr::Float64, small::Float64, smalllpr::Float64, nminpq::Float64
   result = -1.0
   n = k + m
   if (prob > 0.5)
      result = invbinom(k, m, 1.0 - prob, oneMinusP)
   elseif (k = 0.0)
      result = log0(-prob) / n
      if (abs(result) < 1.0)
        result = -expm1(result)
        oneMinusP = 1.0 - result
      else
        oneMinusP = exp(result)
        result = 1.0 - oneMinusP
      end
   elseif (m = 1.0)
      result = log(prob) / n
      if (abs(result) < 1.0)
        oneMinusP = -expm1(result)
        result = 1.0 - oneMinusP
      else
        result = exp(result)
        oneMinusP = 1.0 - result
      end
   else
      pr = -1.0
      xp = invcnormal(prob)
      xp2 = xp * xp
      temp = 2.0 * xp * abs2(k * (m / n) + xp2 / 4.0)
      xp2 = xp2 / n
      dfm = (xp2 * (m - k) + temp) / (2.0 * (1.0 + xp2))
      if (k + dfm < 0.0)
         dfm = -k
      end
      q = (m - dfm) / n
      p = (k + dfm) / n
      dif = -dfm / n
      if (dif = 0.0)
         dif = 1.0
      elseif min(k, m) > 1E+50
         oneMinusP = q
         invcompbinom = p
         Exit Function
      end
      lpr = -log(prob)
      small = 0.00000000000004
      smalllpr = small * lpr * prob
      nminpq = n * min(p, q)
      while ((abs(pr - prob) > smalllpr) && (n * abs(dif) > min(abs(dfm), nminpq) * small))
         pr = compbinomial(k, m, p, q, dfm)
         if (pr < nearly_zero) #/*Should not be happenning often */
            dif = dif / 2.0
            dfm = dfm + n * dif
            p = p + dif
            q = q - dif
            pr = -1.0
         else
            temp2 = binomialTerm(k, m, p, q, dfm, 0.0) * m / q
            if (temp2 < nearly_zero) #/*Should not be happenning often */
               dif = dif / 2.0
               dfm = dfm + n * dif
               p = p + dif
               q = q - dif
               pr = -1.0
            else
               dif = 2.0 * pr * logdif(pr, prob) / (1.0 + abs2(log(prob) / log(pr))) / temp2
               if (q + dif <= 0.0) #/*not v. good */
                  dif = -0.9999 * q
                  dfm = dfm - n * dif
                  p = p - dif
                  q = q + dif
                  pr = -1.0
               elseif (p - dif <= 0.0) #/*v. good */
                  temp = exp(log(prob / pr) / (k + 1.0))
                  dif = p
                  p = temp * p
                  dif = p - dif
                  dfm = n * p - k
                  q = 1.0 - p
                  pr = -1.0
               else
                  dfm = dfm - n * dif
                  p = p - dif
                  q = q + dif
               end
            end
         end
      end
      result = p
      oneMinusP = q
   end
   invcompbinom = result
end

function  abMinuscd(a::Float64, b::Float64, c::Float64, d::Float64)::Float64
   Dim a1::Float64, b1::Float64, c1::Float64, d1::Float64, a2::Float64, b2::Float64, c2::Float64, d2::Float64, r1::Float64, r2::Float64, r2a::Float64, r3::Float64
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
   if (r2a < 0.0) = (r2 < 0.0)
      abMinuscd = (r3 + 2.0 * r2a) + ((r2 - r2a) + r1)
   else
      abMinuscd = r3 + ((r2a + r2) + r1)
   end
end

function  aTimes2Powerb(a::Float64, b As Integer)::Float64
   if b > 709
      a = (a * scalefactor) * scalefactor
      b = b - 512
   elseif b < -709
      a = (a * scalefactor2) * scalefactor2
      b = b + 512
   end
   aTimes2Powerb = a * (2.0) ^ b
end

function  GeneralabMinuscd(a::Float64, b::Float64, c::Float64, d::Float64)::Float64
   Dim s::Float64, ca::Float64, cb::Float64, cc::Float64, cd::Float64
   Dim l2 As Integer, pa As Integer, pb As Integer, pc As Integer, pd As Integer
   s = a * b - c * d
   if a <= 0.0 || b <= 0.0 || c <= 0.0 || d <= 0.0
      GeneralabMinuscd = s
      Exit Function
   elseif s < 0.0
      GeneralabMinuscd = -GeneralabMinuscd(c, d, a, b)
      Exit Function
   end
   l2 = Int(log(a) / log(2.0))
   pa = 51 - l2
   ca = aTimes2Powerb(a, pa)
   l2 = Int(log(b) / log(2.0))
   pb = 51 - l2
   cb = aTimes2Powerb(b, pb)
   l2 = Int(log(c) / log(2.0))
   pc = 51 - l2
   cc = aTimes2Powerb(c, pc)
   pd = pa + pb - pc
   cd = aTimes2Powerb(d, pd)
   GeneralabMinuscd = aTimes2Powerb(abMinuscd(ca, cb, cc, cd), -(pa + pb))
end

function  hypergeometricTerm(ai::Float64, aji::Float64, aki::Float64, amkji::Float64)::Float64
#// Probability that hypergeometric variate from a population with total type Is of aki+ai, total type IIs of amkji+aji, has ai type Is and aji type IIs selected.
   Dim aj::Float64, am::Float64, ak::Float64, amj::Float64, amk::Float64
   Dim cjkmi::Float64, ai1::Float64, aj1::Float64, ak1::Float64, am1::Float64, aki1::Float64, aji1::Float64, amk1::Float64, amj1::Float64, amkji1::Float64
   Dim c1::Float64, c3::Float64, c4::Float64, c5::Float64, loghypergeometricTerm::Float64

   ak = aki + ai
   amk = amkji + aji
   aj = aji + ai
   am = amk + ak
   amj = amkji + aki
   if (am > max_discrete)
      hypergeometricTerm = [#VALUE!]
      Exit Function
   end
   if ((ai = 0.0) && ((aji <= 0.0) || (aki <= 0.0) || (amj < 0.0) || (amk < 0.0)))
      hypergeometricTerm = 1.0
   elseif ((ai > 0.0) && (min(aki, aji) = 0.0) && (max(amj, amk) = 0.0))
      hypergeometricTerm = 1.0
   elseif ((ai >= 0.0) && (amkji > -1.0) && (aki > -1.0) && (aji >= 0.0))
     #c1 = logfbit(amkji) + logfbit(aki) + logfbit(aji) + logfbit(am) + logfbit(ai)
     #c1 = logfbit(amk) + logfbit(ak) + logfbit(aj) + logfbit(amj) - c1
     c1 = lfbaccdif1(ak, amk) - lfbaccdif1(ai, aki) - lfbaccdif1(ai, aji) - lfbaccdif1(aki, amkji) - logfbit(ai)
     ai1 = ai + 1.0
     aj1 = aj + 1.0
     ak1 = ak + 1.0
     am1 = am + 1.0
     aki1 = aki + 1.0
     aji1 = aji + 1.0
     amk1 = amk + 1.0
     amj1 = amj + 1.0
     amkji1 = amkji + 1.0
     cjkmi = GeneralabMinuscd(aji, aki, ai, amkji)
     c5 = (cjkmi - ai) / (amkji1 * am1)
     if (c5 < minLog1Value)
        c3 = amkji * (log((amj1 * amk1) / (amkji1 * am1)) - c5) - c5
     else
        c3 = amkji * log1(c5) - c5
     end

     c5 = (-cjkmi - aji) / (aki1 * am1)
     if (c5 < minLog1Value)
        c4 = aki * (log((ak1 * amj1) / (aki1 * am1)) - c5) - c5
     else
        c4 = aki * log1(c5) - c5
     end

     c3 = c3 + c4
     c5 = (-cjkmi - aki) / (aji1 * am1)
     if (c5 < minLog1Value)
        c4 = aji * (log((aj1 * amk1) / (aji1 * am1)) - c5) - c5
     else
        c4 = aji * log1(c5) - c5
     end

     c3 = c3 + c4
     c5 = (cjkmi - amkji) / (ai1 * am1)
     if (c5 < minLog1Value)
        c4 = ai * (log((aj1 * ak1) / (ai1 * am1)) - c5) - c5
     else
        c4 = ai * log1(c5) - c5
     end

     c3 = c3 + c4
     loghypergeometricTerm = (c1 + 1.0 / am1) + c3

     hypergeometricTerm = exp(loghypergeometricTerm) * abs2((amk1 * ak1) * (aj1 * amj1) / ((amkji1 * aki1 * aji1) * (am1 * ai1))) * OneOverSqrTwoPi
   else
     hypergeometricTerm = 0.0
   end

end

function  hypergeometric(ai::Float64, aji::Float64, aki::Float64, amkji::Float64, comp::Bool, ByRef ha1::Float64, ByRef hprob::Float64, ByRef hswap::Bool)::Float64
#// Probability that hypergeometric variate from a population with total type Is of aki+ai, total type IIs of amkji+aji, has up to ai type Is selected in a sample of size aji+ai.
     Dim prob::Float64
     Dim a1::Float64, a2::Float64, b1::Float64, b2::Float64, an::Float64, bn::Float64, bnAdd::Float64, s::Float64
     Dim c1::Float64, c2::Float64, c3::Float64, c4::Float64
     Dim i::Float64, ji::Float64, ki::Float64, mkji::Float64, njj::Float64, numb::Float64, maxSums::Float64, swapped::Bool
     Dim ip1::Float64, must_do_cf::Bool, allIntegral::Bool, exact::Bool
     if (amkji > -1.0) && (amkji < 0.0)
        ip1 = -amkji
        mkji = ip1 - 1.0
        allIntegral = false
     else
        ip1 = amkji + 1.0
        mkji = amkji
        allIntegral = ai = Int(ai) && aji = Int(aji) && aki = Int(aki) && mkji = Int(mkji)
     end

     if allIntegral
        swapped = (ai + 0.5) * (mkji + 0.5) >= (aki - 0.5) * (aji - 0.5)
     elseif ai < 100.0 && ai = Int(ai) || mkji < 0.0
        if comp
           swapped = (ai + 0.5) * (mkji + 0.5) >= aki * aji
        else
           swapped = (ai + 0.5) * (mkji + 0.5) >= aki * aji + 1000.0
        end
     elseif ai < 1.0
        swapped = (ai + 0.5) * (mkji + 0.5) >= aki * aji
     elseif aji < 1.0 || aki < 1.0 || (ai < 1.0 && ai > 0.0)
        swapped = false
     else
        swapped = (ai + 0.5) * (mkji + 0.5) >= (aki - 0.5) * (aji - 0.5)
     end
     if Not swapped
       i = ai
       ji = aji
       ki = aki
     else
       i = aji - 1.0
       ji = ai + 1.0
       ki = ip1
       ip1 = aki
       mkji = aki - 1.0
     end
     c2 = ji + i
     c4 = mkji + ki + c2
     if (c4 > max_discrete)
        hypergeometric = [#VALUE!]
        Exit Function
     end
     if ((i >= 0.0) && ((ji <= 0.0) || (ki <= 0.0)) || (ip1 + ki <= 0.0) || (ip1 + ji <= 0.0))
        exact = true
        if (i >= 0.0)
           prob = 1.0
        else
           prob = 0.0
        end
     elseif (ip1 > 0.0) && (ip1 < 1.0)
        exact = false
        prob = hypergeometricTerm(i, ji, ki, ip1) * (ip1 * (c4 + 1.0)) / ((ki + ip1) * (ji + ip1))
     else
        exact = ((i = 0.0) && ((ji <= 0.0) || (ki <= 0.0) || (mkji + ki < 0.0) || (mkji + ji < 0.0))) || ((i > 0.0) && (min(ki, ji) = 0.0) && (max(mkji + ki, mkji + ji) = 0.0))
        prob = hypergeometricTerm(i, ji, ki, mkji)
     end
     hprob = prob
     hswap = swapped
     ha1 = 0.0

     if (exact) || (prob = 0.0)
        if (swapped = comp)
           hypergeometric = prob
        else
           hypergeometric = 1.0 - prob
        end
        Exit Function
     end

     a1 = 0.0
     Dim sumAlways As Long, sumFactor As Long
     sumAlways = 0.0
     sumFactor = 10.0

     if i < mkji
        must_do_cf = i <> Int(i)
        maxSums = Int(i)
     else
        must_do_cf = mkji <> Int(mkji)
        maxSums = Int(max(mkji, 0.0))
     end
     if must_do_cf
        sumAlways = 0.0
        sumFactor = 5.0
     else
        sumAlways = 20.0
        sumFactor = 10.0
     end
     if (maxSums > sumAlways || must_do_cf)
        numb = Int(sumFactor / c4 * exp(log((ki + i) * (ji + i) * (ip1 + ji) * (ip1 + ki)) / 3.0))
        numb = Int(i - (ki + i) * (ji + i) / c4 + numb)
        if (numb < 0.0)
           numb = 0.0
        elseif numb > maxSums
           numb = maxSums
        end
     else
        numb = maxSums
     end

     if (2.0 * numb <= maxSums || must_do_cf)
        b1 = 1.0
        c1 = 0.0
        c2 = i - numb
        c3 = mkji - numb
        s = c3
        a2 = c2
        c3 = c3 - 1.0
        b2 = GeneralabMinuscd(ki + numb + 1.0, ji + numb + 1.0, c2 - 1.0, c3)
        bn = b2
        bnAdd = c3 + c4 + c2 - 2.0
        while (b2 > 0.0 && (abs(a2 * b1 - a1 * b2) > abs(cfVSmall * b1 * a2)))
            c1 = c1 + 1.0
            c2 = c2 - 1.0
            an = (c1 * c2) * (c3 * c4)
            c3 = c3 - 1.0
            c4 = c4 - 1.0
            bn = bn + bnAdd
            bnAdd = bnAdd - 4.0
            a1 = bn * a2 + an * a1
            b1 = bn * b2 + an * b1
            if (b1 > scalefactor)
              a1 = a1 * scalefactor2
              b1 = b1 * scalefactor2
              a2 = a2 * scalefactor2
              b2 = b2 * scalefactor2
            end
            c1 = c1 + 1.0
            c2 = c2 - 1.0
            an = (c1 * c2) * (c3 * c4)
            c3 = c3 - 1.0
            c4 = c4 - 1.0
            bn = bn + bnAdd
            bnAdd = bnAdd - 4.0
            a2 = bn * a1 + an * a2
            b2 = bn * b1 + an * b2
            if (b2 > scalefactor)
              a1 = a1 * scalefactor2
              b1 = b1 * scalefactor2
              a2 = a2 * scalefactor2
              b2 = b2 * scalefactor2
            end
        end
        if b1 < 0.0 || b2 < 0.0
           hypergeometric = [#VALUE!]
           Exit Function
        else
           a1 = a2 / b2 * s
        end
     else
        numb = maxSums
     end

     c1 = i - numb + 1.0
     c2 = mkji - numb + 1.0
     c3 = ki + numb
     c4 = ji + numb
     For njj = 1 To numb
       a1 = (1.0 + a1) * ((c1 * c2) / (c3 * c4))
       c1 = c1 + 1.0
       c2 = c2 + 1.0
       c3 = c3 - 1.0
       c4 = c4 - 1.0
     Next njj

     ha1 = a1
     a1 = (1.0 + a1) * prob
     if (swapped = comp)
        hypergeometric = a1
     else
        if a1 > 0.99
           hypergeometric = [#VALUE!]
        else
           hypergeometric = 1.0 - a1
        end
     end
end

function  compgfunc(x::Float64, a::Float64)::Float64
#//Calculates a*x(1/(a+1) - x/2*(1/(a+2) - x/3*(1/(a+3) - ...)))
#//Mainly for calculating the complement of gamma(x,a) for small a and x <= 1.
#//a should be close to 0, x >= 0 & x <=1
  Dim term::Float64, d::Float64, sum::Float64
  term = x
  d = 2.0
  sum = term / (a + 1.0)
  while (abs(term) > abs(sum * sumAcc))
      term = -term * x / d
      sum = sum + term / (a + d)
      d = d + 1.0
  end
  compgfunc = a * sum
end

function  lngammaexpansion(a::Float64)::Float64
#//Calculates log(gamma(a+1)) accurately for for small a (a < 1.5).
#//Uses Abramowitz & Stegun#s series 6.1.33
#//Mainly for calculating the complement of gamma(x,a) for small a and x <= 1.
#//
Dim lgam::Float64
Dim i As Integer
Dim big::Bool
Call initCoeffs
big = a > 0.5
if (big)
   a = a - 1.0
end
i = UBound(coeffs)
lgam = coeffs(i) * logcf(-a / 2.0, i + 2.0, 1.0)
#More accurate with next line for larger values of a
#lgam = logcf(-a / 2.0, i + 2.0, 1.0) * (2.0 ^ (-i - 2)) + logcf(-a / 3.0, i + 2.0, 1.0) * (3.0 ^ (-i - 2))
For i = UBound(coeffs) - 1 To 0 Step -1
   lgam = (coeffs(i) - a * lgam)
Next i
lngammaexpansion = (a * lgam + OneMinusEulers_const) * a
if Not big
   lngammaexpansion = lngammaexpansion - log0(a)
end
end

function  incgamma(x::Float64, a::Float64, comp::Bool)::Float64
#//Calculates gamma-cdf for small a (complementary gamma-cdf if comp).
   Dim r::Float64
   r = a * log(x) - lngammaexpansion(a)
   if (comp)
      r = -expm1(r)
      incgamma = r + compgfunc(x, a) * (1.0 - r)
   else
      incgamma = exp(r) * (1.0 - compgfunc(x, a))
   end
end

function  invincgamma(a::Float64, prob::Float64, comp::Bool)::Float64
#//Calculates inverse of gamma for small a (inverse of complementary gamma if comp).
Dim ga::Float64, x::Float64, deriv::Float64, z::Float64, w::Float64, dif::Float64, pr::Float64, lpr::Float64, small::Float64, smalllpr::Float64
   if (prob > 0.5)
       invincgamma = invincgamma(a, 1.0 - prob, !comp)
       Exit Function
   end
   lpr = -log(prob)
   small = 0.00000000000001
   smalllpr = small * lpr * prob
   if (comp)
      ga = -expm1(lngammaexpansion(a))
      x = -log(prob * (1.0 - ga) / a)
      if (x < 0.5)
         pr = exp(log0(-(ga + prob * (1.0 - ga))) / a)
         if (x < pr)
            x = pr
         end
      end
      dif = x
      pr = -1.0
      while ((abs(pr - prob) > smalllpr) && (abs(dif) > small * max(cSmall, x)))
         deriv = poissonTerm(a, x, x - a, 0.0) * a             #value of derivative is actually deriv/x but it can overflow when x is denormal...
         if (x > 1.0)
            pr = poisson1(-a, x, 0.0)
         else
            z = compgfunc(x, a)
            w = -expm1(a * log(x))
            w = z + w * (1.0 - z)
            pr = (w - ga) / (1.0 - ga)
         end
         dif = x * (pr / deriv) * logdif(pr, prob) #...so multiply by x in slightly different order
         x = x + dif
         if (x < 0.0)
            invincgamma = 0.0
            Exit Function
         end
      end
   else
      ga = exp(lngammaexpansion(a))
      x = log(prob * ga)
      if (x < -711.0 * a)
         invincgamma = 0.0
         Exit Function
      end
      x = exp(x / a)
      z = 1.0 - compgfunc(x, a)
      deriv = poissonTerm(a, x, x - a, 0.0) * a / x
      pr = prob * z
      dif = (pr / deriv) * logdif(pr, prob)
      x = x - dif
      while ((abs(pr - prob) > smalllpr) && (abs(dif) > small * max(cSmall, x)))
         deriv = poissonTerm(a, x, x - a, 0.0) * a / x
         if (x > 1.0)
            pr = 1.0 - poisson1(-a, x, 0.0)
         else
            pr = (1.0 - compgfunc(x, a)) * exp(a * log(x)) / ga
         end
         dif = (pr / deriv) * logdif(pr, prob)
         x = x - dif
      end
   end
   invincgamma = x
end

function  gamma(n::Float64, a::Float64)::Float64
#Assumes n > 0 & a >= 0.  Only called by (comp)gamma_nc with a = 0.
   if (a = 0.0)
      gamma = 1.0
   elseif ((a < 1.0) && (n < 1.0))
      gamma = incgamma(n, a, false)
   elseif (a >= 1.0)
      gamma = comppoisson(a - 1.0, n, n - a + 1.0)
   else
      gamma = 1.0 - poisson1(-a, n, 0.0)
   end
end

function  compgamma(n::Float64, a::Float64)::Float64
#Assumes n > 0 & a >= 0. Only called by (comp)gamma_nc with a = 0.
   if (a = 0.0)
      compgamma = 0.0
   elseif ((a < 1.0) && (n < 1.0))
      compgamma = incgamma(n, a, true)
   elseif (a >= 1.0)
      compgamma = cpoisson(a - 1.0, n, n - a + 1.0)
   else
      compgamma = poisson1(-a, n, 0.0)
   end
end

function  invgamma(a::Float64, prob::Float64)::Float64
#//Inverse of gamma(x,a)
   if (a >= 1.0)
      invgamma = invcomppoisson(a - 1.0, prob)
   else
      invgamma = invincgamma(a, prob, false)
   end
end

function  invcompgamma(a::Float64, prob::Float64)::Float64
#//Inverse of compgamma(x,a)
   if (a >= 1.0)
      invcompgamma = invpoisson(a - 1.0, prob)
   else
      invcompgamma = invincgamma(a, prob, true)
   end
end

function  logfbit1dif(x::Float64)::Float64
#// Calculation of logfbit1(x)-logfbit1(1+x).
  #logfbit1dif = log0(1.0 / (x + 1.0)) - (x + 1.5) / ((x + 1.0) * (x + 2.0))
  logfbit1dif = (logfbitdif(x) - 0.25 / ((x + 1.0) * (x + 2.0))) / (x + 1.5)
end

function  logfbit1(x::Float64)::Float64
#// Derivative of error part of Stirling#s formula where log(x!) = log(sqrt(twopi))+(x+0.5)*log(x+1)-(x+1)+logfbit(x).
  Dim x1::Float64, x2::Float64
  if (x >= 10000000000.0)
     logfbit1 = -lfbc1 * ((x + 1.0) ^ -2)
  elseif (x >= 7.0)
     Dim x3::Float64
     x1 = x + 1.0
     x2 = 1.0 / (x1 * x1)
     x3 = (11.0 * lfbc6 - x2 * (13.0 * lfbc7 - x2 * (15.0 * lfbc8 - x2 * 17.0 * lfbc9)))
     x3 = (5.0 * lfbc3 - x2 * (7.0 * lfbc4 - x2 * (9.0 * lfbc5 - x2 * x3)))
     x3 = x2 * (3.0 * lfbc2 - x2 * x3)
    logfbit1 = -lfbc1 * (1.0 - x3) * x2
  elseif (x > -1.0)
     x1 = x
     x2 = 0.0
     while (x1 < 7.0)
        x2 = x2 + logfbit1dif(x1)
        x1 = x1 + 1.0
     end
     logfbit1 = x2 + logfbit1(x1)
  else
     logfbit1 = -1E+308
  end
end

function  logfbit2dif(x::Float64)::Float64
#// Calculation of logfbit2(x)-logfbit2(1+x).
  logfbit2dif = 0.5 * (((x + 1.0) * (x + 2.0)) ^ -2)
end

function  logfbit2(x::Float64)::Float64
#// Second derivative of error part of Stirling#s formula where log(x!) = log(sqrt(twopi))+(x+0.5)*log(x+1)-(x+1)+logfbit(x).
  Dim x1::Float64, x2::Float64
  if (x >= 10000000000.0)
     logfbit2 = 2.0 * lfbc1 * ((x + 1.0) ^ -3)
  elseif (x >= 7.0)
     Dim x3::Float64
     x1 = x + 1.0
     x2 = 1.0 / (x1 * x1)
     x3 = x2 * (240.0 * lfbc8 - x2 * 306.0 * lfbc9)
     x3 = x2 * (132.0 * lfbc6 - x2 * (182.0 * lfbc7 - x3))
     x3 = x2 * (56.0 * lfbc4 - x2 * (90.0 * lfbc5 - x3))
     x3 = x2 * (12.0 * lfbc2 - x2 * (30.0 * lfbc3 - x3))
     logfbit2 = lfbc1 * (2.0 - x3) * x2 / x1
  elseif (x > -1.0)
     x1 = x
     x2 = 0.0
     while (x1 < 7.0)
        x2 = x2 + logfbit2dif(x1)
        x1 = x1 + 1.0
     end
     logfbit2 = x2 + logfbit2(x1)
  else
     logfbit2 = -1E+308
  end
end

function  logfbit3dif(x::Float64)::Float64
#// Calculation of logfbit3(x)-logfbit3(1+x).
  logfbit3dif = -(2.0 * x + 3.0) * (((x + 1.0) * (x + 2.0)) ^ -3)
end

function  logfbit3(x::Float64)::Float64
#// Third derivative of error part of Stirling#s formula where log(x!) = log(sqrt(twopi))+(x+0.5)*log(x+1)-(x+1)+logfbit(x).
  Dim x1::Float64, x2::Float64
  if (x >= 10000000000.0)
     logfbit3 = -0.5 * ((x + 1.0) ^ -4)
  elseif (x >= 7.0)
     Dim x3::Float64
     x1 = x + 1.0
     x2 = 1.0 / (x1 * x1)
     x3 = x2 * (4080.0 * lfbc8 - x2 * 5814.0 * lfbc9)
     x3 = x2 * (1716.0 * lfbc6 - x2 * (2730.0 * lfbc7 - x3))
     x3 = x2 * (504.0 * lfbc4 - x2 * (990.0 * lfbc5 - x3))
     x3 = x2 * (60.0 * lfbc2 - x2 * (210.0 * lfbc3 - x3))
     logfbit3 = -lfbc1 * (6.0 - x3) * x2 * x2
  elseif (x > -1.0)
     x1 = x
     x2 = 0.0
     while (x1 < 7.0)
        x2 = x2 + logfbit3dif(x1)
        x1 = x1 + 1.0
     end
     logfbit3 = x2 + logfbit3(x1)
  else
     logfbit3 = -1E+308
  end
end

function  logfbit4dif(x::Float64)::Float64
#// Calculation of logfbit4(x)-logfbit4(1+x).
  logfbit4dif = (10.0 * x * (x + 3.0) + 23.0) * (((x + 1.0) * (x + 2.0)) ^ -4)
end

function  logfbit4(x::Float64)::Float64
#// Fourth derivative of error part of Stirling#s formula where log(x!) = log(sqrt(twopi))+(x+0.5)*log(x+1)-(x+1)+logfbit(x).
  Dim x1::Float64, x2::Float64
  if (x >= 10000000000.0)
     logfbit4 = -0.5 * ((x + 1.0) ^ -4)
  elseif (x >= 7.0)
     Dim x3::Float64
     x1 = x + 1.0
     x2 = 1.0 / (x1 * x1)
     x3 = x2 * (73440.0 * lfbc8 - x2 * 116280.0 * lfbc9)
     x3 = x2 * (24024.0 * lfbc6 - x2 * (43680.0 * lfbc7 - x3))
     x3 = x2 * (5040.0 * lfbc4 - x2 * (11880.0 * lfbc5 - x3))
     x3 = x2 * (360.0 * lfbc2 - x2 * (1680.0 * lfbc3 - x3))
     logfbit4 = lfbc1 * (24.0 - x3) * x2 * x2 / x1
  elseif (x > -1.0)
     x1 = x
     x2 = 0.0
     while (x1 < 7.0)
        x2 = x2 + logfbit4dif(x1)
        x1 = x1 + 1.0
     end
     logfbit4 = x2 + logfbit4(x1)
  else
     logfbit4 = -1E+308
  end
end

function  logfbit5dif(x::Float64)::Float64
#// Calculation of logfbit5(x)-logfbit5(1+x).
  logfbit5dif = -6.0 * (2.0 * x + 3.0) * ((5.0 * x + 15.0) * x + 12.0) * (((x + 1.0) * (x + 2.0)) ^ -5)
end

function  logfbit5(x::Float64)::Float64
#// Fifth derivative of error part of Stirling#s formula where log(x!) = log(sqrt(twopi))+(x+0.5)*log(x+1)-(x+1)+logfbit(x).
  Dim x1::Float64, x2::Float64
  if (x >= 10000000000.0)
     logfbit5 = -10.0 * ((x + 1.0) ^ -6)
  elseif (x >= 7.0)
     Dim x3::Float64
     x1 = x + 1.0
     x2 = 1.0 / (x1 * x1)
     x3 = x2 * (1395360.0 * lfbc8 - x2 * 2441880.0 * lfbc9)
     x3 = x2 * (360360.0 * lfbc6 - x2 * (742560.0 * lfbc7 - x3))
     x3 = x2 * (55440.0 * lfbc4 - x2 * (154440.0 * lfbc5 - x3))
     x3 = x2 * (2520.0 * lfbc2 - x2 * (15120.0 * lfbc3 - x3))
     logfbit5 = -lfbc1 * (120.0 - x3) * x2 * x2 * x2
  elseif (x > -1.0)
     x1 = x
     x2 = 0.0
     while (x1 < 7.0)
        x2 = x2 + logfbit5dif(x1)
        x1 = x1 + 1.0
     end
     logfbit5 = x2 + logfbit5(x1)
  else
     logfbit5 = -1E+308
  end
end

function  logfbit7dif(x::Float64)::Float64
#// Calculation of logfbit7(x)-logfbit7(1+x).
  logfbit7dif = -120.0 * (2.0 * x + 3.0) * ((((14.0 * x + 84.0) * x + 196.0) * x + 210.0) * x + 87.0) * (((x + 1.0) * (x + 2.0)) ^ -7)
end

function  logfbit7(x::Float64)::Float64
#// Seventh derivative of error part of Stirling#s formula where log(x!) = log(sqrt(twopi))+(x+0.5)*log(x+1)-(x+1)+logfbit(x).
  Dim x1::Float64, x2::Float64
  if (x >= 10000000000.0)
     logfbit7 = -420.0 * ((x + 1.0) ^ -8)
  elseif (x >= 7.0)
     Dim x3::Float64
     x1 = x + 1.0
     x2 = 1.0 / (x1 * x1)
     x3 = x2 * (586051200.0 * lfbc8 - x2 * 1235591280.0 * lfbc9)
     x3 = x2 * (98017920.0 * lfbc6 - x2 * (253955520.0 * lfbc7 - x3))
     x3 = x2 * (8648640.0 * lfbc4 - x2 * (32432400.0 * lfbc5 - x3))
     x3 = x2 * (181440.0 * lfbc2 - x2 * (1663200.0 * lfbc3 - x3))
     logfbit7 = -lfbc1 * (5040.0 - x3) * x2 * x2 * x2 * x2
  elseif (x > -1.0)
     x1 = x
     x2 = 0.0
     while (x1 < 7.0)
        x2 = x2 + logfbit7dif(x1)
        x1 = x1 + 1.0
     end
     logfbit7 = x2 + logfbit7(x1)
  else
     logfbit7 = -1E+308
  end
end

function  lfbaccdif(a::Float64, b::Float64)::Float64
#// This is now always reasonably accurate, although it is not always required to be so when called from incbeta.
   if (a > 0.025 * (a + b + 1.0))
      lfbaccdif = logfbit(a + b) - logfbit(b)
   else
      Dim a2::Float64, ab2::Float64
      a2 = a * a
      ab2 = a / 2.0 + b
      lfbaccdif = a * (logfbit1(ab2) + a2 / 24.0 * (logfbit3(ab2) + a2 / 80.0 * (logfbit5(ab2) + a2 / 168.0 * logfbit7(ab2))))
   end
end

function  compbfunc(x::Float64, a::Float64, b::Float64)::Float64
#// Calculates a*(b-1)*x(1/(a+1) - (b-2)*x/2*(1/(a+2) - (b-3)*x/3*(1/(a+3) - ...)))
#// Mainly for calculating the complement of beta(x,a,b) for small a and b*x < 1.
#// a should be close to 0, x >= 0 & x <=1 & b*x < 1
  Dim term::Float64, d::Float64, sum::Float64
  term = x
  d = 2.0
  sum = term / (a + 1.0)
  while (abs(term) > abs(sum * sumAcc))
      term = -term * (b - d) * x / d
      sum = sum + term / (a + d)
      d = d + 1.0
  end
  compbfunc = a * (b - 1.0) * sum
end

function  incbeta(x::Float64, a::Float64, b::Float64, comp::Bool)::Float64
#// Calculates beta for small a (complementary beta if comp).
   Dim r::Float64
   if (x > 0.5)
      incbeta = incbeta(1.0 - x, b, a, !comp)
   else
      r = (a + b + 0.5) * log1(a / (1.0 + b)) + a * ((a - 0.5) / (1.0 + b) + log((1.0 + b) * x)) - lfbaccdif1(a, b) - lngammaexpansion(a)
      if (comp)
         r = -expm1(r)
         r = r + compbfunc(x, a, b) * (1.0 - r)
         r = r + (a / (a + b)) * (1.0 - r)
      else
         r = exp(r) * (1.0 - compbfunc(x, a, b)) * (b / (a + b))
      end
      incbeta = r
   end
end

function  beta(x::Float64, a::Float64, b::Float64)::Float64
#//Assumes x >= 0 & a >= 0 & b >= 0. Only called with a = 0 or b = 0 by (comp)beta_nc
   if (a = 0.0 && b = 0.0)
      beta = [#VALUE!]
   elseif (a = 0.0)
      beta = 1.0
   elseif (b = 0.0)
      beta = 0.0
   elseif (x <= 0.0)
      beta = 0.0
   elseif (x >= 1.0)
      beta = 1.0
   elseif (a < 1.0 && b < 1.0)
      beta = incbeta(x, a, b, false)
   elseif (a < 1.0 && (1.0 + b) * x <= 1.0)
      beta = incbeta(x, a, b, false)
   elseif (b < 1.0 && a <= (1.0 + a) * x)
      beta = incbeta(1.0 - x, b, a, true)
   elseif (a < 1.0)
      beta = compbinomial(-a, b, x, 1.0 - x, 0.0)
   elseif (b < 1.0)
      beta = binomial(-b, a, 1.0 - x, x, 0.0)
   else
      beta = compbinomial(a - 1.0, b, x, 1.0 - x, (a + b - 1.0) * x - a + 1.0)
   end
end

function  compbeta(x::Float64, a::Float64, b::Float64)::Float64
#//Assumes x >= 0 & a >= 0 & b >= 0. Only called with a = 0 or b = 0 by (comp)beta_nc
   if (a = 0.0 && b = 0.0)
      compbeta = [#VALUE!]
   elseif (a = 0.0)
      compbeta = 0.0
   elseif (b = 0.0)
      compbeta = 1.0
   elseif (x <= 0.0)
      compbeta = 1.0
   elseif (x >= 1.0)
      compbeta = 0.0
   elseif (a < 1.0 && b < 1.0)
      compbeta = incbeta(x, a, b, true)
   elseif (a < 1.0 && (1.0 + b) * x <= 1.0)
      compbeta = incbeta(x, a, b, true)
   elseif (b < 1.0 && a <= (1.0 + a) * x)
      compbeta = incbeta(1.0 - x, b, a, false)
   elseif (a < 1.0)
      compbeta = binomial(-a, b, x, 1.0 - x, 0.0)
   elseif (b < 1.0)
      compbeta = compbinomial(-b, a, 1.0 - x, x, 0.0)
   else
      compbeta = binomial(a - 1.0, b, x, 1.0 - x, (a + b - 1.0) * x - a + 1.0)
   end
end

function  invincbeta(a::Float64, b::Float64, prob::Float64, comp::Bool, ByRef oneMinusP::Float64)::Float64
#// Calculates inverse of beta for small a (inverse of complementary beta if comp).
Dim r::Float64, rb::Float64, x::Float64, OneOverDeriv::Float64, dif::Float64, pr::Float64, mnab::Float64, aplusbOvermxab::Float64, lpr::Float64, small::Float64, smalllpr::Float64
   if (Not comp && prob > b / (a + b))
       invincbeta = invincbeta(a, b, 1.0 - prob, !comp, oneMinusP)
       Exit Function
   elseif (comp && prob > a / (a + b) && prob > 0.1)
       invincbeta = invincbeta(a, b, 1.0 - prob, !comp, oneMinusP)
       Exit Function
   end
   lpr = max(-log(prob), 1.0)
   small = 0.00000000000001
   smalllpr = small * lpr * prob
   if a >= b
      mnab = b
      aplusbOvermxab = (a + b) / a
   else
      mnab = a
      aplusbOvermxab = (a + b) / b
   end
   if (comp)
      r = (a + b + 0.5) * log1(a / (1.0 + b)) + a * (a - 0.5) / (1.0 + b) - lfbaccdif1(a, b) - lngammaexpansion(a)
      r = -expm1(r)
      r = r + (a / (a + b)) * (1.0 - r)
      if (b < 1.0)
         rb = (a + b + 0.5) * log1(b / (1.0 + a)) + b * (b - 0.5) / (1.0 + a) - lfbaccdif1(b, a) - lngammaexpansion(b)
         rb = exp(rb) * (a / (a + b))
         oneMinusP = log(prob / rb) / b
         if (oneMinusP < 0.0)
             oneMinusP = exp(oneMinusP) / (1.0 + a)
         else
             oneMinusP = 0.5
         end
         if (oneMinusP = 0.0)
            invincbeta = 1.0
            Exit Function
         elseif (oneMinusP > 0.5)
            oneMinusP = 0.5
         end
         x = 1.0 - oneMinusP
         pr = rb * (1.0 - compbfunc(oneMinusP, b, a)) * exp(b * log((1.0 + a) * oneMinusP))
         OneOverDeriv = (aplusbOvermxab * x * oneMinusP) / (binomialTerm(a, b, x, oneMinusP, (a + b) * x - a, 0.0) * mnab)
         dif = OneOverDeriv * pr * logdif(pr, prob)
         oneMinusP = oneMinusP - dif
         x = 1.0 - oneMinusP
         if (oneMinusP <= 0.0)
            oneMinusP = 0.0
            invincbeta = 1.0
            Exit Function
         elseif (x < 0.25)
            x = exp(log0((r - prob) / (1.0 - r)) / a) / (b + 1.0)
            oneMinusP = 1.0 - x
            if (x = 0.0)
               invincbeta = 0.0
               Exit Function
            end
            pr = compbfunc(x, a, b) * (1.0 - prob)
            OneOverDeriv = (aplusbOvermxab * x * oneMinusP) / (binomialTerm(a, b, x, oneMinusP, (a + b) * x - a, 0.0) * mnab)
            dif = OneOverDeriv * (prob + pr) * log0(pr / prob)
            x = x + dif
            if (x <= 0.0)
               oneMinusP = 1.0
               invincbeta = 0.0
               Exit Function
            end
            oneMinusP = 1.0 - x
         end
      else
         pr = exp(log0((r - prob) / (1.0 - r)) / a) / (b + 1.0)
         x = log(b * prob / (a * (1.0 - r) * b * exp(a * log(1.0 + b)))) / b
         if (abs(x) < 0.5)
            x = -expm1(x)
            oneMinusP = 1.0 - x
         else
            oneMinusP = exp(x)
            x = 1.0 - oneMinusP
            if (oneMinusP = 0.0)
               invincbeta = x
               Exit Function
            end
         end
         if pr > x && pr < 1.0
            x = pr
            oneMinusP = 1.0 - x
         end
      end
      dif = min(x, oneMinusP)
      pr = -1.0
      while ((abs(pr - prob) > smalllpr) && (abs(dif) > small * max(cSmall, min(x, oneMinusP))))
         if (b < 1.0 && x > 0.5)
            pr = rb * (1.0 - compbfunc(oneMinusP, b, a)) * exp(b * log((1.0 + a) * oneMinusP))
         elseif ((1.0 + b) * x > 1.0)
            pr = binomial(-a, b, x, oneMinusP, 0.0)
         else
            pr = r + compbfunc(x, a, b) * (1.0 - r)
            pr = pr - expm1(a * log((1.0 + b) * x)) * (1.0 - pr)
         end
         OneOverDeriv = (aplusbOvermxab * x * oneMinusP) / (binomialTerm(a, b, x, oneMinusP, (a + b) * x - a, 0.0) * mnab)
         dif = OneOverDeriv * pr * logdif(pr, prob)
         if (x > 0.5)
            oneMinusP = oneMinusP - dif
            x = 1.0 - oneMinusP
            if (oneMinusP <= 0.0)
               oneMinusP = 0.0
               invincbeta = 1.0
               Exit Function
            end
         else
            x = x + dif
            oneMinusP = 1.0 - x
            if (x <= 0.0)
               oneMinusP = 1.0
               invincbeta = 0.0
               Exit Function
            end
         end
      end
   else
      r = (a + b + 0.5) * log1(a / (1.0 + b)) + a * (a - 0.5) / (1.0 + b) - lfbaccdif1(a, b) - lngammaexpansion(a)
      r = exp(r) * (b / (a + b))
      x = logdif(prob, r)
      if (x < -711.0 * a)
         x = 0.0
      else
         x = exp(x / a) / (1.0 + b)
      end
      if (x = 0.0)
         oneMinusP = 1.0
         invincbeta = x
         Exit Function
      elseif (x >= 0.5)
         x = 0.5
      end
      oneMinusP = 1.0 - x
      pr = r * (1.0 - compbfunc(x, a, b)) * exp(a * log((1.0 + b) * x))
      OneOverDeriv = (aplusbOvermxab * x * oneMinusP) / (binomialTerm(a, b, x, oneMinusP, (a + b) * x - a, 0.0) * mnab)
      dif = OneOverDeriv * pr * logdif(pr, prob)
      x = x - dif
      oneMinusP = oneMinusP + dif
      while ((abs(pr - prob) > smalllpr) && (abs(dif) > small * max(cSmall, min(x, oneMinusP))))
         if ((1.0 + b) * x > 1.0)
            pr = compbinomial(-a, b, x, oneMinusP, 0.0)
         elseif (x > 0.5)
            pr = incbeta(oneMinusP, b, a, !comp)
         else
            pr = r * (1.0 - compbfunc(x, a, b)) * exp(a * log((1.0 + b) * x))
         end
         OneOverDeriv = (aplusbOvermxab * x * oneMinusP) / (binomialTerm(a, b, x, oneMinusP, (a + b) * x - a, 0.0) * mnab)
         dif = OneOverDeriv * pr * logdif(pr, prob)
         if x < 0.5
            x = x - dif
            oneMinusP = 1.0 - x
         else
            oneMinusP = oneMinusP + dif
            x = 1.0 - oneMinusP
         end
      end
   end
   invincbeta = x
end

function  invbeta(a::Float64, b::Float64, prob::Float64, ByRef oneMinusP::Float64)::Float64
   Dim swap::Float64
   if (prob = 0.0)
      oneMinusP = 1.0
      invbeta = 0.0
   elseif (prob = 1.0)
      oneMinusP = 0.0
      invbeta = 1.0
   elseif (a = b && prob = 0.5)
      oneMinusP = 0.5
      invbeta = 0.5
   elseif (a < b && b < 1.0)
      invbeta = invincbeta(a, b, prob, false, oneMinusP)
   elseif (b < a && a < 1.0)
      swap = invincbeta(b, a, prob, true, oneMinusP)
      invbeta = oneMinusP
      oneMinusP = swap
   elseif (a < 1.0)
      invbeta = invincbeta(a, b, prob, false, oneMinusP)
   elseif (b < 1.0)
      swap = invincbeta(b, a, prob, true, oneMinusP)
      invbeta = oneMinusP
      oneMinusP = swap
   else
      invbeta = invcompbinom(a - 1.0, b, prob, oneMinusP)
   end
end

function  invcompbeta(a::Float64, b::Float64, prob::Float64, ByRef oneMinusP::Float64)::Float64
   Dim swap::Float64
   if (prob = 0.0)
      oneMinusP = 0.0
      invcompbeta = 1.0
   elseif (prob = 1.0)
      oneMinusP = 1.0
      invcompbeta = 0.0
   elseif (a = b && prob = 0.5)
      oneMinusP = 0.5
      invcompbeta = 0.5
   elseif (a < b && b < 1.0)
      invcompbeta = invincbeta(a, b, prob, true, oneMinusP)
   elseif (b < a && a < 1.0)
      swap = invincbeta(b, a, prob, false, oneMinusP)
      invcompbeta = oneMinusP
      oneMinusP = swap
   elseif (a < 1.0)
      invcompbeta = invincbeta(a, b, prob, true, oneMinusP)
   elseif (b < 1.0)
      swap = invincbeta(b, a, prob, false, oneMinusP)
      invcompbeta = oneMinusP
      oneMinusP = swap
   else
      invcompbeta = invbinom(a - 1.0, b, prob, oneMinusP)
   end
end

function  critpoiss(mean::Float64, cprob::Float64)::Float64
#//i such that Pr(poisson(mean,i)) >= cprob and  Pr(poisson(mean,i-1)) < cprob
   if (cprob > 0.5)
      critpoiss = critcomppoiss(mean, 1.0 - cprob)
      Exit Function
   end
   Dim pr::Float64, tpr::Float64, dfm::Float64
   Dim i::Float64
   dfm = invcnormal(cprob) * abs2(mean)
   i = Int(mean + dfm + 0.5)
   while (true)
      i = Int(i)
      if (i < 0.0)
         i = 0.0
      end
      if (i >= max_crit)
         critpoiss = i
         Exit Function
      end
      dfm = mean - i
      pr = cpoisson(i, mean, dfm)
      tpr = 0.0
      if (pr >= cprob)
         if (i = 0.0)
            critpoiss = i
            Exit Function
         end
         tpr = poissonTerm(i, mean, dfm, 0.0)
         pr = pr - tpr
         if (pr < cprob)
            critpoiss = i
            Exit Function
         end

         i = i - 1.0
         Dim temp::Float64, temp2::Float64
         temp = (pr - cprob) / tpr
         if (temp > 10)
            temp = Int(temp + 0.5)
            i = i - temp
            temp2 = poissonTerm(i, mean, mean - i, 0.0)
            i = i - temp * (tpr - temp2) / (2 * temp2)
         else
            tpr = tpr * (i + 1.0) / mean
            pr = pr - tpr
            if (pr < cprob)
               critpoiss = i
               Exit Function
            end
            i = i - 1.0
            if (i = 0.0)
               critpoiss = i
               Exit Function
            end
            temp2 = (pr - cprob) / tpr
            if (temp2 < temp - 0.9)
               while (pr >= cprob)
                  tpr = tpr * (i + 1.0) / mean
                  pr = pr - tpr
                  i = i - 1.0
               end
               critpoiss = i + 1.0
               Exit Function
            else
               temp = Int(log(cprob / pr) / log((i + 1.0) / mean) + 0.5)
               i = i - temp
               if (i < 0.0)
                  i = 0.0
               end
               temp2 = poissonTerm(i, mean, mean - i, 0.0)
               if (temp2 > nearly_zero)
                  temp = log((cprob / pr) * (tpr / temp2)) / log((i + 1.0) / mean)
                  i = i - temp
               end
            end
         end
      else
         while ((tpr < cSmall) && (pr < cprob))
            i = i + 1.0
            dfm = dfm - 1.0
            tpr = poissonTerm(i, mean, dfm, 0.0)
            pr = pr + tpr
         end
         while (pr < cprob)
            i = i + 1.0
            tpr = tpr * mean / i
            pr = pr + tpr
         end
         critpoiss = i
         Exit Function
      end
   end
end

function  critcomppoiss(mean::Float64, cprob::Float64)::Float64
#//i such that 1-Pr(poisson(mean,i)) > cprob and  1-Pr(poisson(mean,i-1)) <= cprob
   if (cprob > 0.5)
      critcomppoiss = critpoiss(mean, 1.0 - cprob)
      Exit Function
   end
   Dim pr::Float64, tpr::Float64, dfm::Float64
   Dim i::Float64
   dfm = invcnormal(cprob) * abs2(mean)
   i = Int(mean - dfm + 0.5)
   while (true)
      i = Int(i)
      if (i >= max_crit)
         critcomppoiss = i
         Exit Function
      end
      dfm = mean - i
      pr = comppoisson(i, mean, dfm)
      tpr = 0.0
      if (pr > cprob)
         i = i + 1.0
         dfm = dfm - 1.0
         tpr = poissonTerm(i, mean, dfm, 0.0)
         if (pr < (1.00001) * tpr)
            while (tpr > cprob)
               i = i + 1.0
               tpr = tpr * mean / i
            end
         else
            pr = pr - tpr
            if (pr <= cprob)
               critcomppoiss = i
               Exit Function
            end
            Dim temp::Float64, temp2::Float64
            temp = (pr - cprob) / tpr
            if (temp > 10)
               temp = Int(temp + 0.5)
               i = i + temp
               temp2 = poissonTerm(i, mean, mean - i, 0.0)
               i = i + temp * (tpr - temp2) / (2.0 * temp2)
            elseif (pr / tpr > 0.00001)
               i = i + 1.0
               tpr = tpr * mean / i
               pr = pr - tpr
               if (pr <= cprob)
                  critcomppoiss = i
                  Exit Function
               end
               temp2 = (pr - cprob) / tpr
               if (temp2 < temp - 0.9)
                  while (pr > cprob)
                     i = i + 1.0
                     tpr = tpr * mean / i
                     pr = pr - tpr
                  end
                  critcomppoiss = i
                  Exit Function
               else
                  temp = log(cprob / pr) / log(mean / i)
                  temp = Int((log(cprob / pr) - temp * log(i / (temp + i))) / log(mean / i) + 0.5)
                  i = i + temp
                  temp2 = poissonTerm(i, mean, mean - i, 0.0)
                  if (temp2 > nearly_zero)
                     temp = log((cprob / pr) * (tpr / temp2)) / log(mean / i)
                     i = i + temp
                  end
               end
            end
         end
      else
         while ((tpr < cSmall) && (pr <= cprob))
            tpr = poissonTerm(i, mean, dfm, 0.0)
            pr = pr + tpr
            i = i - 1.0
            dfm = dfm + 1.0
         end
         while (pr <= cprob)
            tpr = tpr * (i + 1.0) / mean
            pr = pr + tpr
            i = i - 1.0
         end
         critcomppoiss = i + 1.0
         Exit Function
      end
   end
end

function  critbinomial(n::Float64, eprob::Float64, cprob::Float64)::Float64
#//i such that Pr(binomial(n,eprob,i)) >= cprob and  Pr(binomial(n,eprob,i-1)) < cprob
   if (cprob > 0.5)
      critbinomial = critcompbinomial(n, eprob, 1.0 - cprob)
      Exit Function
   end
   Dim pr::Float64, tpr::Float64, dfm::Float64
   Dim i::Float64
   dfm = invcnormal(cprob) * abs2(n * eprob * (1.0 - eprob))
   i = n * eprob + dfm
   while (true)
      i = Int(i)
      if (i < 0.0)
         i = 0.0
      elseif (i > n)
         i = n
      end
      if (i >= max_crit)
         critbinomial = i
         Exit Function
      end
      dfm = n * eprob - i
      pr = binomial(i, n - i, eprob, 1.0 - eprob, dfm)
      tpr = 0.0
      if (pr >= cprob)
         if (i = 0.0)
            critbinomial = i
            Exit Function
         end
         tpr = binomialTerm(i, n - i, eprob, 1.0 - eprob, dfm, 0.0)
         if (pr < (1.00001) * tpr)
            tpr = tpr * ((i + 1.0) * (1.0 - eprob)) / ((n - i) * eprob)
            i = i - 1.0
            while (tpr >= cprob)
               tpr = tpr * ((i + 1.0) * (1.0 - eprob)) / ((n - i) * eprob)
               i = i - 1
            end
         else
            pr = pr - tpr
            if (pr < cprob)
               critbinomial = i
               Exit Function
            end
            i = i - 1.0
            if (i = 0.0)
               critbinomial = i
               Exit Function
            end
            Dim temp::Float64, temp2::Float64
            temp = (pr - cprob) / tpr
            if (temp > 10.0)
               temp = Int(temp + 0.5)
               i = i - temp
               temp2 = binomialTerm(i, n - i, eprob, 1.0 - eprob, n * eprob - i, 0.0)
               i = i - temp * (tpr - temp2) / (2.0 * temp2)
            else
               tpr = tpr * ((i + 1.0) * (1.0 - eprob)) / ((n - i) * eprob)
               pr = pr - tpr
               if (pr < cprob)
                  critbinomial = i
                  Exit Function
               end
               i = i - 1.0
               temp2 = (pr - cprob) / tpr
               if (temp2 < temp - 0.9)
                  while (pr >= cprob)
                     tpr = tpr * ((i + 1.0) * (1.0 - eprob)) / ((n - i) * eprob)
                     pr = pr - tpr
                     i = i - 1.0
                  end
                  critbinomial = i + 1.0
                  Exit Function
               else
                  temp = Int(log(cprob / pr) / log(((i + 1.0) * (1.0 - eprob)) / ((n - i) * eprob)) + 0.5)
                  i = i - temp
                  if (i < 0.0)
                     i = 0.0
                  end
                  temp2 = binomialTerm(i, n - i, eprob, 1.0 - eprob, n * eprob - i, 0.0)
                  if (temp2 > nearly_zero)
                     temp = log((cprob / pr) * (tpr / temp2)) / log(((i + 1.0) * (1.0 - eprob)) / ((n - i) * eprob))
                     i = i - temp
                  end
               end
            end
         end
      else
         while ((tpr < cSmall) && (pr < cprob))
            i = i + 1.0
            dfm = dfm - 1.0
            tpr = binomialTerm(i, n - i, eprob, 1.0 - eprob, dfm, 0.0)
            pr = pr + tpr
         end
         while (pr < cprob)
            i = i + 1.0
            tpr = tpr * ((n - i + 1.0) * eprob) / (i * (1.0 - eprob))
            pr = pr + tpr
         end
         critbinomial = i
         Exit Function
      end
   end
end

function  critcompbinomial(n::Float64, eprob::Float64, cprob::Float64)::Float64
#//i such that 1-Pr(binomial(n,eprob,i)) > cprob and  1-Pr(binomial(n,eprob,i-1)) <= cprob
   if (cprob > 0.5)
      critcompbinomial = critbinomial(n, eprob, 1.0 - cprob)
      Exit Function
   end
   Dim pr::Float64, tpr::Float64, dfm::Float64
   Dim i::Float64
   dfm = invcnormal(cprob) * abs2(n * eprob * (1.0 - eprob))
   i = n * eprob - dfm
   while (true)
      i = Int(i)
      if (i < 0.0)
         i = 0.0
      elseif (i > n)
         i = n
      end
      if (i >= max_crit)
         critcompbinomial = i
         Exit Function
      end
      dfm = n * eprob - i
      pr = compbinomial(i, n - i, eprob, 1.0 - eprob, dfm)
      tpr = 0.0
      if (pr > cprob)
         i = i + 1.0
         dfm = dfm - 1.0
         tpr = binomialTerm(i, n - i, eprob, 1.0 - eprob, dfm, 0.0)
         if (pr < (1.00001) * tpr)
            while (tpr > cprob)
               i = i + 1.0
               tpr = tpr * ((n - i + 1.0) * eprob) / (i * (1.0 - eprob))
            end
         else
            pr = pr - tpr
            if (pr <= cprob)
               critcompbinomial = i
               Exit Function
            end
            Dim temp::Float64, temp2::Float64
            temp = (pr - cprob) / tpr
            if (temp > 10.0)
               temp = Int(temp + 0.5)
               i = i + temp
               temp2 = binomialTerm(i, n - i, eprob, 1.0 - eprob, n * eprob - i, 0.0)
               i = i + temp * (tpr - temp2) / (2.0 * temp2)
            else
               i = i + 1.0
               tpr = tpr * ((n - i + 1.0) * eprob) / (i * (1.0 - eprob))
               pr = pr - tpr
               if (pr <= cprob)
                  critcompbinomial = i
                  Exit Function
               end
               temp2 = (pr - cprob) / tpr
               if (temp2 < temp - 0.9)
                  while (pr > cprob)
                     i = i + 1.0
                     tpr = tpr * ((n - i + 1.0) * eprob) / (i * (1.0 - eprob))
                     pr = pr - tpr
                  end
                  critcompbinomial = i
                  Exit Function
               else
                  temp = Int(log(cprob / pr) / log(((n - i + 1.0) * eprob) / (i * (1.0 - eprob))) + 0.5)
                  i = i + temp
                  if (i > n)
                     i = n
                  end
                  temp2 = binomialTerm(i, n - i, eprob, 1.0 - eprob, n * eprob - i, 0.0)
                  if (temp2 > nearly_zero)
                     temp = log((cprob / pr) * (tpr / temp2)) / log(((n - i + 1.0) * eprob) / (i * (1.0 - eprob)))
                     i = i + temp
                  end
               end
            end
         end
      else
         while ((tpr < cSmall) && (pr <= cprob))
            tpr = binomialTerm(i, n - i, eprob, 1.0 - eprob, dfm, 0.0)
            pr = pr + tpr
            i = i - 1.0
            dfm = dfm + 1.0
         end
         while (pr <= cprob)
            tpr = tpr * ((i + 1.0) * (1.0 - eprob)) / ((n - i) * eprob)
            pr = pr + tpr
            i = i - 1.0
         end
         critcompbinomial = i + 1.0
         Exit Function
      end
   end
end

function  crithyperg(j::Float64, k::Float64, m::Float64, cprob::Float64)::Float64
#//i such that Pr(hypergeometric(i,j,k,m)) >= cprob and  Pr(hypergeometric(i-1,j,k,m)) < cprob
   Dim ha1::Float64, hprob::Float64, hswap::Bool
   if (cprob > 0.5)
      crithyperg = critcomphyperg(j, k, m, 1.0 - cprob)
      Exit Function
   end
   Dim pr::Float64, tpr::Float64
   Dim i::Float64
   i = j * k / m + invcnormal(cprob) * abs2(j * k * (m - j) * (m - k) / (m * m * (m - 1.0)))
   Dim mx::Float64, mn ::Float64
   mx = min(j, k)
   mn = max(0, j + k - m)
   while (true)
      if (i < mn)
         i = mn
      elseif (i > mx)
         i = mx
      end
      i = Int(i + 0.5)
      if (i >= max_crit)
         crithyperg = i
         Exit Function
      end
      pr = hypergeometric(i, j - i, k - i, m - k - j + i, false, ha1, hprob, hswap)
      tpr = 0.0
      if (pr >= cprob)
         if (i = mn)
            crithyperg = mn
            Exit Function
         end
         tpr = hypergeometricTerm(i, j - i, k - i, m - k - j + i)
         if (pr < (1.00001) * tpr)
            tpr = tpr * ((i + 1.0) * (m - j - k + i + 1.0)) / ((k - i) * (j - i))
            i = i - 1.0
            while (tpr > cprob)
               tpr = tpr * ((i + 1.0) * (m - j - k + i + 1.0)) / ((k - i) * (j - i))
               i = i - 1.0
            end
         else
            pr = pr - tpr
            if (pr < cprob)
               crithyperg = i
               Exit Function
            end
            i = i - 1.0
            if (i = mn)
               crithyperg = mn
               Exit Function
            end
            Dim temp::Float64, temp2::Float64
            temp = (pr - cprob) / tpr
            if (temp > 10)
               temp = Int(temp + 0.5)
               i = i - temp
               temp2 = hypergeometricTerm(i, j - i, k - i, m - k - j + i)
               i = i - temp * (tpr - temp2) / (2.0 * temp2)
            else
               tpr = tpr * ((i + 1.0) * (m - j - k + i + 1.0)) / ((k - i) * (j - i))
               pr = pr - tpr
               if (pr < cprob)
                  crithyperg = i
                  Exit Function
               end
               i = i - 1.0
               temp2 = (pr - cprob) / tpr
               if (temp2 < temp - 0.9)
                  while (pr >= cprob)
                     tpr = tpr * ((i + 1.0) * (m - j - k + i + 1.0)) / ((k - i) * (j - i))
                     pr = pr - tpr
                     i = i - 1.0
                  end
                  crithyperg = i + 1.0
                  Exit Function
               else
                  temp = Int(log(cprob / pr) / log(((i + 1.0) * (m - j - k + i + 1.0)) / ((k - i) * (j - i))) + 0.5)
                  i = i - temp
                  temp2 = hypergeometricTerm(i, j - i, k - i, m - k - j + i)
                  if (temp2 > nearly_zero)
                     temp = log((cprob / pr) * (tpr / temp2)) / log(((i + 1.0) * (m - j - k + i + 1.0)) / ((k - i) * (j - i)))
                     i = i - temp
                  end
               end
            end
         end
      else
         while ((tpr < cSmall) && (pr < cprob))
            i = i + 1.0
            tpr = hypergeometricTerm(i, j - i, k - i, m - k - j + i)
            pr = pr + tpr
         end
         while (pr < cprob)
            i = i + 1.0
            tpr = tpr * ((k - i + 1.0) * (j - i + 1.0)) / (i * (m - j - k + i))
            pr = pr + tpr
         end
         crithyperg = i
         Exit Function
      end
   end
end

function  critcomphyperg(j::Float64, k::Float64, m::Float64, cprob::Float64)::Float64
#//i such that 1-Pr(hypergeometric(i,j,k,m)) > cprob and  1-Pr(hypergeometric(i-1,j,k,m)) <= cprob
   Dim ha1::Float64, hprob::Float64, hswap::Bool
   if (cprob > 0.5)
      critcomphyperg = crithyperg(j, k, m, 1.0 - cprob)
      Exit Function
   end
   Dim pr::Float64, tpr::Float64
   Dim i::Float64
   i = j * k / m - invcnormal(cprob) * abs2(j * k * (m - j) * (m - k) / (m * m * (m - 1.0)))
   Dim mx::Float64, mn ::Float64
   mx = min(j, k)
   mn = max(0, j + k - m)
   while (true)
      if (i < mn)
         i = mn
      elseif (i > mx)
         i = mx
      end
      i = Int(i + 0.5)
      if (i >= max_crit)
         critcomphyperg = i
         Exit Function
      end
      pr = hypergeometric(i, j - i, k - i, m - k - j + i, true, ha1, hprob, hswap)
      tpr = 0.0
      if (pr > cprob)
         i = i + 1.0
         tpr = hypergeometricTerm(i, j - i, k - i, m - k - j + i)
         if (pr < (1.0 + 0.00001) * tpr)
            while (tpr > cprob)
               i = i + 1
               tpr = tpr * ((k - i + 1.0) * (j - i + 1.0)) / (i * (m - j - k + i))
            end
         else
            pr = pr - tpr
            if (pr <= cprob)
               critcomphyperg = i
               Exit Function
            end
            Dim temp::Float64, temp2::Float64
            temp = (pr - cprob) / tpr
            if (temp > 10)
               temp = Int(temp + 0.5)
               i = i + temp
               temp2 = hypergeometricTerm(i, j - i, k - i, m - k - j + i)
               i = i + temp * (tpr - temp2) / (2.0 * temp2)
            else
               i = i + 1.0
               tpr = tpr * ((k - i + 1.0) * (j - i + 1.0)) / (i * (m - j - k + i))
               pr = pr - tpr
               if (pr <= cprob)
                  critcomphyperg = i
                  Exit Function
               end
               temp2 = (pr - cprob) / tpr
               if (temp2 < temp - 0.9)
                  while (pr > cprob)
                     i = i + 1.0
                     tpr = tpr * ((k - i + 1.0) * (j - i + 1.0)) / (i * (m - j - k + i))
                     pr = pr - tpr
                  end
                  critcomphyperg = i
                  Exit Function
               else
                  temp = Int(log(cprob / pr) / log(((k - i + 1.0) * (j - i + 1.0)) / (i * (m - j - k + i))) + 0.5)
                  i = i + temp
                  temp2 = hypergeometricTerm(i, j - i, k, m - k)
                  if (temp2 > nearly_zero)
                     temp = log((cprob / pr) * (tpr / temp2)) / log(((k - i + 1.0) * (j - i + 1.0)) / (i * (m - j - k + i)))
                     i = i + temp
                  end
               end
            end
         end
      else
         while ((tpr < cSmall) && (pr <= cprob))
            tpr = hypergeometricTerm(i, j - i, k - i, m - k - j + i)
            pr = pr + tpr
            i = i - 1.0
         end
         while (pr <= cprob)
            tpr = tpr * ((i + 1.0) * (m - j - k + i + 1.0)) / ((k - i) * (j - i))
            pr = pr + tpr
            i = i - 1.0
         end
         critcomphyperg = i + 1.0
         Exit Function
      end
   end
end

function  critnegbinom(n::Float64, eprob::Float64, fprob::Float64, cprob::Float64)::Float64
#//i such that Pr(negbinomial(n,eprob,i)) >= cprob and  Pr(negbinomial(n,eprob,i-1)) < cprob
   if (cprob > 0.5)
      critnegbinom = critcompnegbinom(n, eprob, fprob, 1.0 - cprob)
      Exit Function
   end
   Dim pr::Float64, tpr::Float64, dfm::Float64
   Dim i::Float64
   i = invgamma(n * fprob, cprob) / eprob
   while (true)
      if (i < 0.0)
         i = 0.0
      end
      i = Int(i)
      if (i >= max_crit)
         critnegbinom = i
         Exit Function
      end
      if eprob <= fprob
         pr = beta(eprob, n, i + 1.0)
      else
         pr = compbeta(fprob, i + 1.0, n)
      end
      tpr = 0.0
      if (pr >= cprob)
         if (i = 0.0)
            critnegbinom = i
            Exit Function
         end
         if eprob <= fprob
            dfm = n - (n + i) * eprob
         else
            dfm = (n + i) * fprob - i
         end
         tpr = n / (n + i) * binomialTerm(i, n, fprob, eprob, dfm, 0.0)
         if (pr < (1.00001) * tpr)
            tpr = tpr * (i + 1.0) / ((n + i) * fprob)
            i = i - 1.0
            while (tpr > cprob)
               tpr = tpr * (i + 1.0) / ((n + i) * fprob)
               i = i - 1.0
            end
         else
            pr = pr - tpr
            if (pr < cprob)
               critnegbinom = i
               Exit Function
            end
            i = i - 1.0
            if (i = 0.0)
               critnegbinom = i
               Exit Function
            end
            Dim temp::Float64, temp2::Float64
            temp = (pr - cprob) / tpr
            if (temp > 10.0)
               temp = Int(temp + 0.5)
               i = i - temp
               if eprob <= fprob
                  dfm = n - (n + i) * eprob
               else
                  dfm = (n + i) * fprob - i
               end
               temp2 = n / (n + i) * binomialTerm(i, n, fprob, eprob, dfm, 0.0)
               i = i - temp * (tpr - temp2) / (2.0 * temp2)
            else
               tpr = tpr * (i + 1.0) / ((n + i) * fprob)
               pr = pr - tpr
               if (pr < cprob)
                  critnegbinom = i
                  Exit Function
               end
               i = i - 1.0
               temp2 = (pr - cprob) / tpr
               if (temp2 < temp - 0.9)
                  while (pr >= cprob)
                     tpr = tpr * (i + 1.0) / ((n + i) * fprob)
                     pr = pr - tpr
                     i = i - 1.0
                  end
                  critnegbinom = i + 1.0
                  Exit Function
               else
                  temp = Int(log(cprob / pr) / log((i + 1.0) / ((n + i) * fprob)) + 0.5)
                  i = i - temp
                  if eprob <= fprob
                     dfm = n - (n + i) * eprob
                  else
                     dfm = (n + i) * fprob - i
                  end
                  temp2 = n / (n + i) * binomialTerm(i, n, fprob, eprob, dfm, 0.0)
                  if (temp2 > nearly_zero)
                     temp = log((cprob / pr) * (tpr / temp2)) / log((i + 1.0) / ((n + i) * fprob))
                     i = i - temp
                  end
               end
            end
         end
      else
         while ((tpr < cSmall) && (pr < cprob))
            i = i + 1.0
            if eprob <= fprob
               dfm = n - (n + i) * eprob
            else
               dfm = (n + i) * fprob - i
            end
            tpr = n / (n + i) * binomialTerm(i, n, fprob, eprob, dfm, 0.0)
            pr = pr + tpr
         end
         while (pr < cprob)
            i = i + 1.0
            tpr = tpr * ((n + i - 1.0) * fprob) / i
            pr = pr + tpr
         end
         critnegbinom = i
         Exit Function
      end
   end
end

function  critcompnegbinom(n::Float64, eprob::Float64, fprob::Float64, cprob::Float64)::Float64
#//i such that 1-Pr(negbinomial(n,eprob,i)) > cprob and  1-Pr(negbinomial(n,eprob,i-1)) <= cprob
   if (cprob > 0.5)
      critcompnegbinom = critnegbinom(n, eprob, fprob, 1.0 - cprob)
      Exit Function
   end
   Dim pr::Float64, tpr::Float64, dfm::Float64
   Dim i::Float64
   i = invcompgamma(n * fprob, cprob) / eprob
   while (true)
      if (i < 0.0)
         i = 0.0
      end
      i = Int(i)
      if (i >= max_crit)
         critcompnegbinom = i
         Exit Function
      end
      if eprob <= fprob
         pr = compbeta(eprob, n, i + 1.0)
      else
         pr = beta(fprob, i + 1.0, n)
      end
      if (pr > cprob)
         i = i + 1.0
         if eprob <= fprob
            dfm = n - (n + i) * eprob
         else
            dfm = (n + i) * fprob - i
         end
         tpr = n / (n + i) * binomialTerm(i, n, fprob, eprob, dfm, 0.0)
         if (pr < (1.00001) * tpr)
            while (tpr > cprob)
               i = i + 1.0
               tpr = tpr * ((n + i - 1.0) * fprob) / i
            end
         else
            pr = pr - tpr
            if (pr <= cprob)
               critcompnegbinom = i
               Exit Function
            elseif (tpr < 0.000000000000001 * pr)
               if (tpr < cSmall)
                  critcompnegbinom = i
               else
                  critcompnegbinom = i + Int((pr - cprob) / tpr)
               end
               Exit Function
            end
            Dim temp::Float64, temp2::Float64
            temp = (pr - cprob) / tpr
            if (temp > 10.0)
               temp = Int(temp + 0.5)
               i = i + temp
               if eprob <= fprob
                  dfm = n - (n + i) * eprob
               else
                  dfm = (n + i) * fprob - i
               end
               temp2 = n / (n + i) * binomialTerm(i, n, fprob, eprob, dfm, 0.0)
               i = i + temp * (tpr - temp2) / (2.0 * temp2)
            else
               i = i + 1.0
               tpr = tpr * ((n + i - 1.0) * fprob) / i
               pr = pr - tpr
               if (pr <= cprob)
                  critcompnegbinom = i
                  Exit Function
               end
               temp2 = (pr - cprob) / tpr
               if (temp2 < temp - 0.9)
                  while (pr > cprob)
                     i = i + 1.0
                     tpr = tpr * ((n + i - 1.0) * fprob) / i
                     pr = pr - tpr
                  end
                  critcompnegbinom = i
                  Exit Function
               else
                  temp = Int(log(cprob / pr) / log(((n + i - 1.0) * fprob) / i) + 0.5)
                  i = i + temp
                  if eprob <= fprob
                     dfm = n - (n + i) * eprob
                  else
                     dfm = (n + i) * fprob - i
                  end
                  temp2 = n / (n + i) * binomialTerm(i, n, fprob, eprob, dfm, 0.0)
                  if (temp2 > nearly_zero)
                     temp = log((cprob / pr) * (tpr / temp2)) / log(((n + i - 1.0) * fprob) / i)
                     i = i + temp
                  end
               end
            end
         end
      else
         if eprob <= fprob
            dfm = n - (n + i) * eprob
         else
            dfm = (n + i) * fprob - i
         end
         tpr = n / (n + i) * binomialTerm(i, n, fprob, eprob, dfm, 0.0)
         if (tpr < 0.000000000000001 * pr)
            if (tpr < cSmall)
               critcompnegbinom = i
            else
               critcompnegbinom = i - Int((cprob - pr) / tpr)
            end
            Exit Function
         end
         while ((tpr < cSmall) && (pr <= cprob))
            pr = pr + tpr
            i = i - 1.0
            if eprob <= fprob
               dfm = n - (n + i) * eprob
            else
               dfm = (n + i) * fprob - i
            end
            tpr = n / (n + i) * binomialTerm(i, n, fprob, eprob, dfm, 0.0)
         end
         while (pr <= cprob)
            pr = pr + tpr
            i = i - 1.0
            if i < 0.0
               critcompnegbinom = 0.0
               Exit Function
            end
            tpr = tpr * (i + 1.0) / ((n + i) * fprob)
         end
         critcompnegbinom = i + 1.0
         Exit Function
      end
   end
end

function  critneghyperg(j::Float64, k::Float64, m::Float64, cprob::Float64)::Float64
#//i such that Pr(neghypergeometric(i,j,k,m)) >= cprob and  Pr(neghypergeometric(i-1,j,k,m)) < cprob
   Dim ha1::Float64, hprob::Float64, hswap::Bool
   if (cprob > 0.5)
      critneghyperg = critcompneghyperg(j, k, m, 1.0 - cprob)
      Exit Function
   end
   Dim pr::Float64, tpr::Float64
   Dim i::Float64, temp::Float64, temp2::Float64, oneMinusP::Float64
   pr = (m - k) / m
   i = invbeta(j * pr, pr * (k - j + 1.0), cprob, oneMinusP) * (m - k)
   while (true)
      if (i < 0.0)
         i = 0.0
      elseif (i > m - k)
         i = m - k
      end
      i = Int(i + 0.5)
      if (i >= max_crit)
         critneghyperg = i
         Exit Function
      end
      pr = hypergeometric(i, j, m - k - i, k - j, false, ha1, hprob, hswap)
      tpr = 0.0
      if (pr >= cprob)
         if (i = 0.0)
            critneghyperg = 0.0
            Exit Function
         end
         tpr = hypergeometricTerm(j - 1.0, i, k - j + 1.0, m - k - i) * (k - j + 1.0) / (m - j - i + 1.0)
         if (pr < (1.0 + 0.00001) * tpr)
            tpr = tpr * ((i + 1.0) * (m - j - i)) / ((j + i) * (m - i - k))
            i = i - 1.0
            while (tpr > cprob)
               tpr = tpr * ((i + 1.0) * (m - j - i)) / ((j + i) * (m - i - k))
               i = i - 1.0
            end
         else
            pr = pr - tpr
            if (pr < cprob)
               critneghyperg = i
               Exit Function
            end
            i = i - 1.0

            if (i = 0.0)
               critneghyperg = 0.0
               Exit Function
            end
            temp = (pr - cprob) / tpr
            if (temp > 10)
               temp = Int(temp + 0.5)
               i = i - temp
               temp2 = hypergeometricTerm(j - 1.0, i, k - j + 1.0, m - k - i) * (k - j + 1.0) / (m - j - i + 1.0)
               i = i - temp * (tpr - temp2) / (2 * temp2)
            else
               tpr = tpr * ((i + 1.0) * (m - j - i)) / ((j + i) * (m - i - k))
               pr = pr - tpr
               if (pr < cprob)
                  critneghyperg = i
                  Exit Function
               end
               i = i - 1.0
               temp2 = (pr - cprob) / tpr
               if (temp2 < temp - 0.9)
                  while (pr >= cprob)
                     tpr = tpr * ((i + 1.0) * (m - j - i)) / ((j + i) * (m - i - k))
                     pr = pr - tpr
                     i = i - 1.0
                  end
                  critneghyperg = i + 1.0
                  Exit Function
               else
                  temp = Int(log(cprob / pr) / log(((i + 1.0) * (m - j - i)) / ((j + i) * (m - i - k))) + 0.5)
                  i = i - temp
                  temp2 = hypergeometricTerm(j - 1.0, i, k - j + 1.0, m - k - i) * (k - j + 1.0) / (m - j - i + 1.0)
                  if (temp2 > nearly_zero)
                     temp = log((cprob / pr) * (tpr / temp2)) / log(((i + 1.0) * (m - j - i)) / ((j + i) * (m - i - k)))
                     i = i - temp
                  end
               end
            end
         end
      else
         while ((tpr < cSmall) && (pr < cprob))
            i = i + 1.0
            tpr = hypergeometricTerm(j - 1.0, i, k - j + 1.0, m - k - i) * (k - j + 1.0) / (m - j - i + 1.0)
            pr = pr + tpr
         end
         while (pr < cprob)
            i = i + 1.0
            tpr = tpr * ((j + i - 1.0) * (m - i - k + 1.0)) / (i * (m - j - i + 1.0))
            pr = pr + tpr
         end
         critneghyperg = i
         Exit Function
      end
   end
end

function  critcompneghyperg(j::Float64, k::Float64, m::Float64, cprob::Float64)::Float64
#//i such that 1-Pr(neghypergeometric(i,j,k,m)) > cprob and  1-Pr(neghypergeometric(i-1,j,k,m)) <= cprob
   Dim ha1::Float64, hprob::Float64, hswap::Bool
   if (cprob > 0.5)
      critcompneghyperg = critneghyperg(j, k, m, 1.0 - cprob)
      Exit Function
   end
   Dim pr::Float64, tpr::Float64
   Dim i::Float64, temp::Float64, temp2::Float64, oneMinusP::Float64
   pr = (m - k) / m
   i = invcompbeta(j * pr, pr * (k - j + 1.0), cprob, oneMinusP) * (m - k)
   while (true)
      if (i < 0.0)
         i = 0.0
      elseif (i > m - k)
         i = m - k
      end
      i = Int(i + 0.5)
      if (i >= max_crit)
         critcompneghyperg = i
         Exit Function
      end
      pr = hypergeometric(i, j, m - k - i, k - j, true, ha1, hprob, hswap)
      tpr = 0.0
      if (pr > cprob)
         i = i + 1.0
         tpr = hypergeometricTerm(j - 1.0, i, k - j + 1.0, m - k - i) * (k - j + 1.0) / (m - j - i + 1.0)
         if (pr < (1.00001) * tpr)
            while (tpr > cprob)
               i = i + 1.0
               temp = m - j - i + 1.0
               if temp = 0.0 Exit Do
               tpr = tpr * ((j + i - 1.0) * (m - i - k + 1.0)) / (i * temp)
            end
         else
            pr = pr - tpr
            if (pr <= cprob)
               critcompneghyperg = i
               Exit Function
            end
            temp = (pr - cprob) / tpr
            if (temp > 10.0)
               temp = Int(temp + 0.5)
               i = i + temp
               temp2 = hypergeometricTerm(j - 1.0, i, k - j + 1.0, m - k - i) * (k - j + 1.0) / (m - j - i + 1.0)
               i = i + temp * (tpr - temp2) / (2 * temp2)
            else
               i = i + 1.0
               tpr = tpr * ((j + i - 1.0) * (m - i - k + 1.0)) / (i * (m - j - i + 1.0))
               pr = pr - tpr
               if (pr <= cprob)
                  critcompneghyperg = i
                  Exit Function
               end
               temp2 = (pr - cprob) / tpr
               if (temp2 < temp - 0.9)
                  while (pr > cprob)
                     i = i + 1.0
                     tpr = tpr * ((j + i - 1.0) * (m - i - k + 1.0)) / (i * (m - j - i + 1.0))
                     pr = pr - tpr
                  end
                  critcompneghyperg = i
                  Exit Function
               else
                  temp = Int(log(cprob / pr) / log(((j + i - 1.0) * (m - i - k + 1.0)) / (i * (m - j - i + 1.0))) + 0.5)
                  i = i + temp
                  temp2 = hypergeometricTerm(j - 1.0, i, k - j + 1.0, m - k - i) * (k - j + 1.0) / (m - j - i + 1.0)
                  if (temp2 > nearly_zero)
                     temp = log((cprob / pr) * (tpr / temp2)) / log(((j + i - 1.0) * (m - i - k + 1.0)) / (i * (m - j - i + 1.0)))
                     i = i + temp
                  end
               end
            end
         end
      else
         while ((tpr < cSmall) && (pr <= cprob))
            tpr = hypergeometricTerm(j - 1.0, i, k - j + 1.0, m - k - i) * (k - j + 1.0) / (m - j - i + 1.0)
            pr = pr + tpr
            i = i - 1.0
         end
         while (pr <= cprob)
            tpr = tpr * ((i + 1.0) * (m - j - i)) / ((j + i) * (m - i - k))
            pr = pr + tpr
            i = i - 1.0
         end
         critcompneghyperg = i + 1.0
         Exit Function
      end
   end
end

function  AlterForIntegralChecks_Others(value::Float64)::Float64
   if NonIntegralValuesAllowed_Others
      AlterForIntegralChecks_Others = Int(value)
   elseif value <> Int(value)
      AlterForIntegralChecks_Others = [#VALUE!]
   else
      AlterForIntegralChecks_Others = value
   end
end

function  AlterForIntegralChecks_df(value::Float64)::Float64
   if NonIntegralValuesAllowed_df
      AlterForIntegralChecks_df = value
   else
      AlterForIntegralChecks_df = AlterForIntegralChecks_Others(value)
   end
end

function  AlterForIntegralChecks_NB(value::Float64)::Float64
   if NonIntegralValuesAllowed_NB
      AlterForIntegralChecks_NB = value
   else
      AlterForIntegralChecks_NB = AlterForIntegralChecks_Others(value)
   end
end

function  GetRidOfMinusZeroes(x::Float64)::Float64
   if x = 0.0
      GetRidOfMinusZeroes = 0.0
   else
      GetRidOfMinusZeroes = x
   end
end

function pmf_geometric(failures::Float64, success_prob::Float64)::Float64
   failures = AlterForIntegralChecks_Others(failures)
   if (success_prob < 0.0 || success_prob > 1.0)
      pmf_geometric = [#VALUE!]
   elseif failures < 0.0
      pmf_geometric = 0.0
   elseif success_prob = 1.0
      if failures = 0.0
         pmf_geometric = 1.0
      else
         pmf_geometric = 0.0
      end
   else
      pmf_geometric = success_prob * exp(failures * log0(-success_prob))
   end
   pmf_geometric = GetRidOfMinusZeroes(pmf_geometric)
end

function cdf_geometric(failures::Float64, success_prob::Float64)::Float64
   failures = Int(failures)
   if (success_prob < 0.0 || success_prob > 1.0)
      cdf_geometric = [#VALUE!]
   elseif failures < 0.0
      cdf_geometric = 0.0
   elseif success_prob = 1.0
      if failures >= 0.0
         cdf_geometric = 1.0
      else
         cdf_geometric = 0.0
      end
   else
      cdf_geometric = -expm1((failures + 1.0) * log0(-success_prob))
   end
   cdf_geometric = GetRidOfMinusZeroes(cdf_geometric)
end

function comp_cdf_geometric(failures::Float64, success_prob::Float64)::Float64
   failures = Int(failures)
   if (success_prob < 0.0 || success_prob > 1.0)
      comp_cdf_geometric = [#VALUE!]
   elseif failures < 0.0
      comp_cdf_geometric = 1.0
   elseif success_prob = 1.0
      if failures >= 0.0
         comp_cdf_geometric = 0.0
      else
         comp_cdf_geometric = 1.0
      end
   else
      comp_cdf_geometric = exp((failures + 1.0) * log0(-success_prob))
   end
   comp_cdf_geometric = GetRidOfMinusZeroes(comp_cdf_geometric)
end

function crit_geometric(success_prob::Float64, crit_prob::Float64)::Float64
   if (success_prob <= 0.0 || success_prob > 1.0 || crit_prob < 0.0 || crit_prob > 1.0)
      crit_geometric = [#VALUE!]
   elseif (crit_prob = 0.0)
      crit_geometric = [#VALUE!]
   elseif (success_prob = 1.0)
      crit_geometric = 0.0
   elseif (crit_prob = 1.0)
      crit_geometric = [#VALUE!]
   else
      crit_geometric = Int(log0(-crit_prob) / log0(-success_prob) - 1.0)
      if -expm1((crit_geometric + 1.0) * log0(-success_prob)) < crit_prob
         crit_geometric = crit_geometric + 1.0
      end
   end
   crit_geometric = GetRidOfMinusZeroes(crit_geometric)
end

function comp_crit_geometric(success_prob::Float64, crit_prob::Float64)::Float64
   if (success_prob <= 0.0 || success_prob > 1.0 || crit_prob < 0.0 || crit_prob > 1.0)
      comp_crit_geometric = [#VALUE!]
   elseif (crit_prob = 1.0)
      comp_crit_geometric = [#VALUE!]
   elseif (success_prob = 1.0)
      comp_crit_geometric = 0.0
   elseif (crit_prob = 0.0)
      comp_crit_geometric = [#VALUE!]
   else
      comp_crit_geometric = Int(log(crit_prob) / log0(-success_prob) - 1.0)
      if exp((comp_crit_geometric + 1.0) * log0(-success_prob)) > crit_prob
         comp_crit_geometric = comp_crit_geometric + 1.0
      end
   end
   comp_crit_geometric = GetRidOfMinusZeroes(comp_crit_geometric)
end

function lcb_geometric(failures::Float64, prob::Float64)::Float64
   failures = AlterForIntegralChecks_Others(failures)
   if (prob < 0.0 || prob > 1.0 || failures < 0.0)
      lcb_geometric = [#VALUE!]
   elseif (prob = 1.0)
      lcb_geometric = 1.0
   else
      lcb_geometric = -expm1(log0(-prob) / (failures + 1.0))
   end
   lcb_geometric = GetRidOfMinusZeroes(lcb_geometric)
end

function ucb_geometric(failures::Float64, prob::Float64)::Float64
   failures = AlterForIntegralChecks_Others(failures)
   if (prob < 0.0 || prob > 1.0 || failures < 0.0)
      ucb_geometric = [#VALUE!]
   elseif (prob = 0.0 || failures = 0.0)
      ucb_geometric = 1.0
   elseif (prob = 1.0)
      ucb_geometric = 0.0
   else
      ucb_geometric = -expm1(log(prob) / failures)
   end
   ucb_geometric = GetRidOfMinusZeroes(ucb_geometric)
end

function pmf_negbinomial(failures::Float64, success_prob::Float64, successes_reqd::Float64)::Float64
   Dim q::Float64, dfm::Float64
   failures = AlterForIntegralChecks_Others(failures)
   successes_reqd = AlterForIntegralChecks_NB(successes_reqd)
   if (success_prob < 0.0 || success_prob > 1.0 || successes_reqd <= 0.0)
      pmf_negbinomial = [#VALUE!]
   elseif (successes_reqd + failures > 0.0)
      q = 1.0 - success_prob
      if success_prob <= q
         dfm = successes_reqd - (successes_reqd + failures) * success_prob
      else
         dfm = (successes_reqd + failures) * q - failures
      end
      pmf_negbinomial = successes_reqd / (successes_reqd + failures) * binomialTerm(failures, successes_reqd, q, success_prob, dfm, 0.0)
   elseif (failures <> 0.0)
      pmf_negbinomial = 0.0
   else
      pmf_negbinomial = 1.0
   end
   pmf_negbinomial = GetRidOfMinusZeroes(pmf_negbinomial)
end

function cdf_negbinomial(failures::Float64, success_prob::Float64, successes_reqd::Float64)::Float64
   Dim q::Float64, dfm::Float64
   failures = Int(failures)
   successes_reqd = AlterForIntegralChecks_NB(successes_reqd)
   if (success_prob < 0.0 || success_prob > 1.0 || successes_reqd <= 0.0)
      cdf_negbinomial = [#VALUE!]
   else
      q = 1.0 - success_prob
      if q < success_prob
         cdf_negbinomial = compbeta(q, failures + 1, successes_reqd)
      else
         cdf_negbinomial = beta(success_prob, successes_reqd, failures + 1)
      end
   end
   cdf_negbinomial = GetRidOfMinusZeroes(cdf_negbinomial)
end

function comp_cdf_negbinomial(failures::Float64, success_prob::Float64, successes_reqd::Float64)::Float64
   Dim q::Float64, dfm::Float64
   failures = Int(failures)
   successes_reqd = AlterForIntegralChecks_NB(successes_reqd)
   if (success_prob < 0.0 || success_prob > 1.0 || successes_reqd <= 0.0)
      comp_cdf_negbinomial = [#VALUE!]
   else
      q = 1.0 - success_prob
      if q < success_prob
         comp_cdf_negbinomial = beta(q, failures + 1, successes_reqd)
      else
         comp_cdf_negbinomial = compbeta(success_prob, successes_reqd, failures + 1)
      end
   end
   comp_cdf_negbinomial = GetRidOfMinusZeroes(comp_cdf_negbinomial)
end

function crit_negbinomial(success_prob::Float64, successes_reqd::Float64, crit_prob::Float64)::Float64
   successes_reqd = AlterForIntegralChecks_NB(successes_reqd)
   if (success_prob <= 0.0 || success_prob > 1.0 || successes_reqd <= 0.0 || crit_prob < 0.0 || crit_prob > 1.0)
      crit_negbinomial = [#VALUE!]
   elseif (crit_prob = 0.0)
      crit_negbinomial = [#VALUE!]
   elseif (success_prob = 1.0)
      crit_negbinomial = 0.0
   elseif (crit_prob = 1.0)
      crit_negbinomial = [#VALUE!]
   else
      Dim i::Float64, pr::Float64
      crit_negbinomial = critnegbinom(successes_reqd, success_prob, 1.0 - success_prob, crit_prob)
      i = crit_negbinomial
      pr = cdf_negbinomial(i, success_prob, successes_reqd)
      if (pr = crit_prob)
      elseif (pr > crit_prob)
         i = i - 1.0
         pr = cdf_negbinomial(i, success_prob, successes_reqd)
         if (pr >= crit_prob)
            crit_negbinomial = i
         end
      else
         crit_negbinomial = i + 1.0
      end
   end
   crit_negbinomial = GetRidOfMinusZeroes(crit_negbinomial)
end

function comp_crit_negbinomial(success_prob::Float64, successes_reqd::Float64, crit_prob::Float64)::Float64
   successes_reqd = AlterForIntegralChecks_NB(successes_reqd)
   if (success_prob <= 0.0 || success_prob > 1.0 || successes_reqd <= 0.0 || crit_prob < 0.0 || crit_prob > 1.0)
      comp_crit_negbinomial = [#VALUE!]
   elseif (crit_prob = 1.0)
      comp_crit_negbinomial = [#VALUE!]
   elseif (success_prob = 1.0)
      comp_crit_negbinomial = 0.0
   elseif (crit_prob = 0.0)
      comp_crit_negbinomial = [#VALUE!]
   else
      Dim i::Float64, pr::Float64
      comp_crit_negbinomial = critcompnegbinom(successes_reqd, success_prob, 1.0 - success_prob, crit_prob)
      i = comp_crit_negbinomial
      pr = comp_cdf_negbinomial(i, success_prob, successes_reqd)
      if (pr = crit_prob)
      elseif (pr < crit_prob)
         i = i - 1.0
         pr = comp_cdf_negbinomial(i, success_prob, successes_reqd)
         if (pr <= crit_prob)
            comp_crit_negbinomial = i
         end
      else
         comp_crit_negbinomial = i + 1.0
      end
   end
   comp_crit_negbinomial = GetRidOfMinusZeroes(comp_crit_negbinomial)
end

function lcb_negbinomial(failures::Float64, successes_reqd::Float64, prob::Float64)::Float64
   failures = AlterForIntegralChecks_Others(failures)
   successes_reqd = AlterForIntegralChecks_NB(successes_reqd)
   if (prob < 0.0 || prob > 1.0 || failures < 0.0 || successes_reqd <= 0.0)
      lcb_negbinomial = [#VALUE!]
   elseif (prob = 0.0)
      lcb_negbinomial = 0.0
   elseif (prob = 1.0)
      lcb_negbinomial = 1.0
   else
      Dim oneMinusP::Float64
      lcb_negbinomial = invbeta(successes_reqd, failures + 1, prob, oneMinusP)
   end
   lcb_negbinomial = GetRidOfMinusZeroes(lcb_negbinomial)
end

function ucb_negbinomial(failures::Float64, successes_reqd::Float64, prob::Float64)::Float64
   failures = AlterForIntegralChecks_Others(failures)
   successes_reqd = AlterForIntegralChecks_NB(successes_reqd)
   if (prob < 0.0 || prob > 1.0 || failures < 0.0 || successes_reqd <= 0.0)
      ucb_negbinomial = [#VALUE!]
   elseif (prob = 0.0 || failures = 0.0)
      ucb_negbinomial = 1.0
   elseif (prob = 1.0)
      ucb_negbinomial = 0.0
   else
      Dim oneMinusP::Float64
      ucb_negbinomial = invcompbeta(successes_reqd, failures, prob, oneMinusP)
   end
   ucb_negbinomial = GetRidOfMinusZeroes(ucb_negbinomial)
end

function pmf_binomial(sample_size::Float64, successes::Float64, success_prob::Float64)::Float64
   Dim q::Float64, dfm::Float64
   successes = AlterForIntegralChecks_Others(successes)
   sample_size = AlterForIntegralChecks_Others(sample_size)
   if (success_prob < 0.0 || success_prob > 1.0 || sample_size < 0.0)
      pmf_binomial = [#VALUE!]
   else
      q = 1.0 - success_prob
      if success_prob <= q
         dfm = sample_size * success_prob - successes
      else
         dfm = (sample_size - successes) - sample_size * q
      end
      pmf_binomial = binomialTerm(successes, sample_size - successes, success_prob, q, dfm, 0.0)
   end
   pmf_binomial = GetRidOfMinusZeroes(pmf_binomial)
end

function cdf_binomial(sample_size::Float64, successes::Float64, success_prob::Float64)::Float64
   Dim q::Float64, dfm::Float64
   successes = Int(successes)
   sample_size = AlterForIntegralChecks_Others(sample_size)
   if (success_prob < 0.0 || success_prob > 1.0 || sample_size < 0.0)
      cdf_binomial = [#VALUE!]
   else
      q = 1.0 - success_prob
      if success_prob <= q
         dfm = sample_size * success_prob - successes
      else
         dfm = (sample_size - successes) - sample_size * q
      end
      cdf_binomial = binomial(successes, sample_size - successes, success_prob, q, dfm)
   end
   cdf_binomial = GetRidOfMinusZeroes(cdf_binomial)
end

function comp_cdf_binomial(sample_size::Float64, successes::Float64, success_prob::Float64)::Float64
   Dim q::Float64, dfm::Float64
   successes = Int(successes)
   sample_size = AlterForIntegralChecks_Others(sample_size)
   if (success_prob < 0.0 || success_prob > 1.0 || sample_size < 0.0)
      comp_cdf_binomial = [#VALUE!]
   else
      q = 1.0 - success_prob
      if success_prob <= q
         dfm = sample_size * success_prob - successes
      else
         dfm = (sample_size - successes) - sample_size * q
      end
      comp_cdf_binomial = compbinomial(successes, sample_size - successes, success_prob, q, dfm)
   end
   comp_cdf_binomial = GetRidOfMinusZeroes(comp_cdf_binomial)
end

function crit_binomial(sample_size::Float64, success_prob::Float64, crit_prob::Float64)::Float64
   sample_size = AlterForIntegralChecks_Others(sample_size)
   if (success_prob < 0.0 || success_prob > 1.0 || sample_size < 0.0 || crit_prob < 0.0 || crit_prob > 1.0)
      crit_binomial = [#VALUE!]
   elseif (crit_prob = 0.0)
      crit_binomial = [#VALUE!]
   elseif (success_prob = 0.0)
      crit_binomial = 0.0
   elseif (crit_prob = 1.0 || success_prob = 1.0)
      crit_binomial = sample_size
   else
      Dim pr::Float64, i::Float64
      crit_binomial = critbinomial(sample_size, success_prob, crit_prob)
      i = crit_binomial
      pr = cdf_binomial(sample_size, i, success_prob)
      if (pr = crit_prob)
      elseif (pr > crit_prob)
         i = i - 1.0
         pr = cdf_binomial(sample_size, i, success_prob)
         if (pr >= crit_prob)
            crit_binomial = i
         end
      else
         crit_binomial = i + 1.0
      end
   end
   crit_binomial = GetRidOfMinusZeroes(crit_binomial)
end

function comp_crit_binomial(sample_size::Float64, success_prob::Float64, crit_prob::Float64)::Float64
   sample_size = AlterForIntegralChecks_Others(sample_size)
   if (success_prob < 0.0 || success_prob > 1.0 || sample_size < 0.0 || crit_prob < 0.0 || crit_prob > 1.0)
      comp_crit_binomial = [#VALUE!]
   elseif (crit_prob = 1.0)
      comp_crit_binomial = [#VALUE!]
   elseif (crit_prob = 0.0 || success_prob = 1.0)
      comp_crit_binomial = sample_size
   elseif (success_prob = 0.0)
      comp_crit_binomial = 0.0
   else
      Dim pr::Float64, i::Float64
      comp_crit_binomial = critcompbinomial(sample_size, success_prob, crit_prob)
      i = comp_crit_binomial
      pr = comp_cdf_binomial(sample_size, i, success_prob)
      if (pr = crit_prob)
      elseif (pr < crit_prob)
         i = i - 1.0
         pr = comp_cdf_binomial(sample_size, i, success_prob)
         if (pr <= crit_prob)
            comp_crit_binomial = i
         end
      else
         comp_crit_binomial = i + 1.0
      end
   end
   comp_crit_binomial = GetRidOfMinusZeroes(comp_crit_binomial)
end

function lcb_binomial(sample_size::Float64, successes::Float64, prob::Float64)::Float64
   successes = AlterForIntegralChecks_Others(successes)
   sample_size = AlterForIntegralChecks_Others(sample_size)
   if (prob < 0.0 || prob > 1.0)
      lcb_binomial = [#VALUE!]
   elseif (sample_size < successes || successes < 0.0)
      lcb_binomial = [#VALUE!]
   elseif (prob = 0.0 || successes = 0.0)
      lcb_binomial = 0.0
   elseif (prob = 1.0)
      lcb_binomial = 1.0
   else
      Dim oneMinusP::Float64
      lcb_binomial = invcompbinom(successes - 1.0, sample_size - successes + 1.0, prob, oneMinusP)
   end
   lcb_binomial = GetRidOfMinusZeroes(lcb_binomial)
end

function ucb_binomial(sample_size::Float64, successes::Float64, prob::Float64)::Float64
   successes = AlterForIntegralChecks_Others(successes)
   sample_size = AlterForIntegralChecks_Others(sample_size)
   if (prob < 0.0 || prob > 1.0)
      ucb_binomial = [#VALUE!]
   elseif (sample_size < successes || successes < 0.0)
      ucb_binomial = [#VALUE!]
   elseif (prob = 0.0 || successes = sample_size#)
      ucb_binomial = 1.0
   elseif (prob = 1.0)
      ucb_binomial = 0.0
   else
      Dim oneMinusP::Float64
      ucb_binomial = invbinom(successes, sample_size - successes, prob, oneMinusP)
   end
   ucb_binomial = GetRidOfMinusZeroes(ucb_binomial)
end

function pmf_poisson(mean::Float64, i::Float64)::Float64
   i = AlterForIntegralChecks_Others(i)
   if (mean < 0.0)
      pmf_poisson = [#VALUE!]
   elseif (i < 0.0)
      pmf_poisson = 0.0
   else
      pmf_poisson = poissonTerm(i, mean, mean - i, 0.0)
   end
   pmf_poisson = GetRidOfMinusZeroes(pmf_poisson)
end

function cdf_poisson(mean::Float64, i::Float64)::Float64
   i = Int(i)
   if (mean < 0.0)
      cdf_poisson = [#VALUE!]
   elseif (i < 0.0)
      cdf_poisson = 0.0
   else
      cdf_poisson = cpoisson(i, mean, mean - i)
   end
   cdf_poisson = GetRidOfMinusZeroes(cdf_poisson)
end

function comp_cdf_poisson(mean::Float64, i::Float64)::Float64
   i = Int(i)
   if (mean < 0.0)
      comp_cdf_poisson = [#VALUE!]
   elseif (i < 0.0)
      comp_cdf_poisson = 1.0
   else
      comp_cdf_poisson = comppoisson(i, mean, mean - i)
   end
   comp_cdf_poisson = GetRidOfMinusZeroes(comp_cdf_poisson)
end

function crit_poisson(mean::Float64, crit_prob::Float64)::Float64
   if (crit_prob < 0.0 || crit_prob > 1.0 || mean < 0.0)
      crit_poisson = [#VALUE!]
   elseif (crit_prob = 0.0)
      crit_poisson = [#VALUE!]
   elseif (mean = 0.0)
      crit_poisson = 0.0
   elseif (crit_prob = 1.0)
      crit_poisson = [#VALUE!]
   else
      Dim pr::Float64
      crit_poisson = critpoiss(mean, crit_prob)
      pr = cpoisson(crit_poisson, mean, mean - crit_poisson)
      if (pr = crit_prob)
      elseif (pr > crit_prob)
         crit_poisson = crit_poisson - 1.0
         pr = cpoisson(crit_poisson, mean, mean - crit_poisson)
         if (pr < crit_prob)
            crit_poisson = crit_poisson + 1.0
         end
      else
         crit_poisson = crit_poisson + 1.0
      end
   end
   crit_poisson = GetRidOfMinusZeroes(crit_poisson)
end

function comp_crit_poisson(mean::Float64, crit_prob::Float64)::Float64
   if (crit_prob < 0.0 || crit_prob > 1.0 || mean < 0.0)
      comp_crit_poisson = [#VALUE!]
   elseif (crit_prob = 1.0)
      comp_crit_poisson = [#VALUE!]
   elseif (mean = 0.0)
      comp_crit_poisson = 0.0
   elseif (crit_prob = 0.0)
      comp_crit_poisson = [#VALUE!]
   else
      Dim pr::Float64
      comp_crit_poisson = critcomppoiss(mean, crit_prob)
      pr = comppoisson(comp_crit_poisson, mean, mean - comp_crit_poisson)
      if (pr = crit_prob)
      elseif (pr < crit_prob)
         comp_crit_poisson = comp_crit_poisson - 1.0
         pr = comppoisson(comp_crit_poisson, mean, mean - comp_crit_poisson)
         if (pr > crit_prob)
            comp_crit_poisson = comp_crit_poisson + 1.0
         end
      else
         comp_crit_poisson = comp_crit_poisson + 1.0
      end
   end
   comp_crit_poisson = GetRidOfMinusZeroes(comp_crit_poisson)
end

function lcb_poisson(i::Float64, prob::Float64)::Float64
   i = AlterForIntegralChecks_Others(i)
   if (prob < 0.0 || prob > 1.0 || i < 0.0)
      lcb_poisson = [#VALUE!]
   elseif (prob = 0.0 || i = 0.0)
      lcb_poisson = 0.0
   elseif (prob = 1.0)
      lcb_poisson = [#VALUE!]
   else
      lcb_poisson = invcomppoisson(i - 1.0, prob)
   end
   lcb_poisson = GetRidOfMinusZeroes(lcb_poisson)
end

function ucb_poisson(i::Float64, prob::Float64)::Float64
   i = AlterForIntegralChecks_Others(i)
   if (prob <= 0.0 || prob > 1.0)
      ucb_poisson = [#VALUE!]
   elseif (i < 0.0)
      ucb_poisson = [#VALUE!]
   elseif (prob = 1.0)
      ucb_poisson = 0.0
   else
      ucb_poisson = invpoisson(i, prob)
   end
   ucb_poisson = GetRidOfMinusZeroes(ucb_poisson)
end

function pmf_hypergeometric(type1s::Float64, sample_size::Float64, tot_type1::Float64, pop_size::Float64)::Float64
   type1s = AlterForIntegralChecks_Others(type1s)
   sample_size = AlterForIntegralChecks_Others(sample_size)
   tot_type1 = AlterForIntegralChecks_Others(tot_type1)
   pop_size = AlterForIntegralChecks_Others(pop_size)
   if (sample_size < 0.0 || tot_type1 < 0.0 || sample_size > pop_size || tot_type1 > pop_size)
      pmf_hypergeometric = [#VALUE!]
   else
      pmf_hypergeometric = hypergeometricTerm(type1s, sample_size - type1s, tot_type1 - type1s, pop_size - tot_type1 - sample_size + type1s)
   end
   pmf_hypergeometric = GetRidOfMinusZeroes(pmf_hypergeometric)
end

function cdf_hypergeometric(type1s::Float64, sample_size::Float64, tot_type1::Float64, pop_size::Float64)::Float64
   type1s = Int(type1s)
   sample_size = AlterForIntegralChecks_Others(sample_size)
   tot_type1 = AlterForIntegralChecks_Others(tot_type1)
   pop_size = AlterForIntegralChecks_Others(pop_size)
   if (sample_size < 0.0 || tot_type1 < 0.0 || sample_size > pop_size || tot_type1 > pop_size)
      cdf_hypergeometric = [#VALUE!]
   else
      Dim ha1::Float64, hprob::Float64, hswap::Bool
      cdf_hypergeometric = hypergeometric(type1s, sample_size - type1s, tot_type1 - type1s, pop_size - tot_type1 - sample_size + type1s, false, ha1, hprob, hswap)
   end
   cdf_hypergeometric = GetRidOfMinusZeroes(cdf_hypergeometric)
end

function comp_cdf_hypergeometric(type1s::Float64, sample_size::Float64, tot_type1::Float64, pop_size::Float64)::Float64
   type1s = Int(type1s)
   sample_size = AlterForIntegralChecks_Others(sample_size)
   tot_type1 = AlterForIntegralChecks_Others(tot_type1)
   pop_size = AlterForIntegralChecks_Others(pop_size)
   if (sample_size < 0.0 || tot_type1 < 0.0 || sample_size > pop_size || tot_type1 > pop_size)
      comp_cdf_hypergeometric = [#VALUE!]
   else
      Dim ha1::Float64, hprob::Float64, hswap::Bool
      comp_cdf_hypergeometric = hypergeometric(type1s, sample_size - type1s, tot_type1 - type1s, pop_size - tot_type1 - sample_size + type1s, true, ha1, hprob, hswap)
   end
   comp_cdf_hypergeometric = GetRidOfMinusZeroes(comp_cdf_hypergeometric)
end

function crit_hypergeometric(sample_size::Float64, tot_type1::Float64, pop_size::Float64, crit_prob::Float64)::Float64
   sample_size = AlterForIntegralChecks_Others(sample_size)
   tot_type1 = AlterForIntegralChecks_Others(tot_type1)
   pop_size = AlterForIntegralChecks_Others(pop_size)
   if (crit_prob < 0.0 || crit_prob > 1.0)
      crit_hypergeometric = [#VALUE!]
   elseif (sample_size < 0.0 || tot_type1 < 0.0 || sample_size > pop_size || tot_type1 > pop_size)
      crit_hypergeometric = [#VALUE!]
   elseif (crit_prob = 0.0)
      crit_hypergeometric = [#VALUE!]
   elseif (sample_size = 0.0 || tot_type1 = 0.0)
      crit_hypergeometric = 0.0
   elseif (sample_size = pop_size || tot_type1 = pop_size)
      crit_hypergeometric = min(sample_size, tot_type1)
   elseif (crit_prob = 1.0)
      crit_hypergeometric = min(sample_size, tot_type1)
   else
      Dim ha1::Float64, hprob::Float64, hswap::Bool
      Dim i::Float64, pr::Float64
      crit_hypergeometric = crithyperg(sample_size, tot_type1, pop_size, crit_prob)
      i = crit_hypergeometric
      pr = hypergeometric(i, sample_size - i, tot_type1 - i, pop_size - tot_type1 - sample_size + i, false, ha1, hprob, hswap)
      if (pr = crit_prob)
      elseif (pr > crit_prob)
         i = i - 1.0
         pr = hypergeometric(i, sample_size - i, tot_type1 - i, pop_size - tot_type1 - sample_size + i, false, ha1, hprob, hswap)
         if (pr >= crit_prob)
            crit_hypergeometric = i
         end
      else
         crit_hypergeometric = i + 1.0
      end
   end
   crit_hypergeometric = GetRidOfMinusZeroes(crit_hypergeometric)
end

function comp_crit_hypergeometric(sample_size::Float64, tot_type1::Float64, pop_size::Float64, crit_prob::Float64)::Float64
   sample_size = AlterForIntegralChecks_Others(sample_size)
   tot_type1 = AlterForIntegralChecks_Others(tot_type1)
   pop_size = AlterForIntegralChecks_Others(pop_size)
   if (crit_prob < 0.0 || crit_prob > 1.0)
      comp_crit_hypergeometric = [#VALUE!]
   elseif (sample_size < 0.0 || tot_type1 < 0.0 || sample_size > pop_size || tot_type1 > pop_size)
      comp_crit_hypergeometric = [#VALUE!]
   elseif (crit_prob = 1.0)
      comp_crit_hypergeometric = [#VALUE!]
   elseif (sample_size = 0.0 || tot_type1 = 0.0)
      comp_crit_hypergeometric = 0.0
   elseif (sample_size = pop_size || tot_type1 = pop_size)
      comp_crit_hypergeometric = min(sample_size, tot_type1)
   elseif (crit_prob = 0.0)
      comp_crit_hypergeometric = min(sample_size, tot_type1)
   else
      Dim ha1::Float64, hprob::Float64, hswap::Bool
      Dim i::Float64, pr::Float64
      comp_crit_hypergeometric = critcomphyperg(sample_size, tot_type1, pop_size, crit_prob)
      i = comp_crit_hypergeometric
      pr = hypergeometric(i, sample_size - i, tot_type1 - i, pop_size - tot_type1 - sample_size + i, true, ha1, hprob, hswap)
      if (pr = crit_prob)
      elseif (pr < crit_prob)
         i = i - 1.0
         pr = hypergeometric(i, sample_size - i, tot_type1 - i, pop_size - tot_type1 - sample_size + i, true, ha1, hprob, hswap)
         if (pr <= crit_prob)
            comp_crit_hypergeometric = i
         end
      else
         comp_crit_hypergeometric = i + 1.0
      end
   end
   comp_crit_hypergeometric = GetRidOfMinusZeroes(comp_crit_hypergeometric)
end

function lcb_hypergeometric(type1s::Float64, sample_size::Float64, pop_size::Float64, prob::Float64)::Float64
   type1s = AlterForIntegralChecks_Others(type1s)
   sample_size = AlterForIntegralChecks_Others(sample_size)
   pop_size = AlterForIntegralChecks_Others(pop_size)
   if (prob < 0.0 || prob > 1.0)
      lcb_hypergeometric = [#VALUE!]
   elseif (type1s < 0.0 || type1s > sample_size || sample_size > pop_size)
      lcb_hypergeometric = [#VALUE!]
   elseif (prob = 0.0 || type1s = 0.0 || pop_size = sample_size)
      lcb_hypergeometric = type1s
   elseif (prob = 1.0)
      lcb_hypergeometric = pop_size - (sample_size - type1s)
   elseif (prob < 0.5)
      lcb_hypergeometric = critneghyperg(type1s, sample_size, pop_size, prob * (1.000000000001)) + type1s
   else
      lcb_hypergeometric = critcompneghyperg(type1s, sample_size, pop_size, (1.0 - prob) * (1.0 - 0.000000000001)) + type1s
   end
   lcb_hypergeometric = GetRidOfMinusZeroes(lcb_hypergeometric)
end

function ucb_hypergeometric(type1s::Float64, sample_size::Float64, pop_size::Float64, prob::Float64)::Float64
   type1s = AlterForIntegralChecks_Others(type1s)
   sample_size = AlterForIntegralChecks_Others(sample_size)
   pop_size = AlterForIntegralChecks_Others(pop_size)
   if (prob < 0.0 || prob > 1.0)
      ucb_hypergeometric = [#VALUE!]
   elseif (type1s < 0.0 || type1s > sample_size || sample_size > pop_size)
      ucb_hypergeometric = [#VALUE!]
   elseif (prob = 0.0 || type1s = sample_size || pop_size = sample_size)
      ucb_hypergeometric = pop_size - (sample_size - type1s)
   elseif (prob = 1.0)
      ucb_hypergeometric = type1s
   elseif (prob < 0.5)
      ucb_hypergeometric = critcompneghyperg(type1s + 1.0, sample_size, pop_size, prob * (1.0 - 0.000000000001)) + type1s
   else
      ucb_hypergeometric = critneghyperg(type1s + 1.0, sample_size, pop_size, (1.0 - prob) * (1.000000000001)) + type1s
   end
   ucb_hypergeometric = GetRidOfMinusZeroes(ucb_hypergeometric)
end

function pmf_neghypergeometric(type2s::Float64, type1s_reqd::Float64, tot_type1::Float64, pop_size::Float64)::Float64
   type2s = AlterForIntegralChecks_Others(type2s)
   type1s_reqd = AlterForIntegralChecks_Others(type1s_reqd)
   tot_type1 = AlterForIntegralChecks_Others(tot_type1)
   pop_size = AlterForIntegralChecks_Others(pop_size)
   if (type1s_reqd <= 0.0 || tot_type1 < type1s_reqd || tot_type1 > pop_size)
      pmf_neghypergeometric = [#VALUE!]
   elseif (type2s < 0.0 || tot_type1 + type2s > pop_size)
      if type2s = 0.0
         pmf_neghypergeometric = 1.0
      else
         pmf_neghypergeometric = 0.0
      end
   else
      pmf_neghypergeometric = hypergeometricTerm(type1s_reqd - 1.0, type2s, tot_type1 - type1s_reqd + 1.0, pop_size - tot_type1 - type2s) * (tot_type1 - type1s_reqd + 1.0) / (pop_size - type1s_reqd - type2s + 1.0)
   end
   pmf_neghypergeometric = GetRidOfMinusZeroes(pmf_neghypergeometric)
end

function cdf_neghypergeometric(type2s::Float64, type1s_reqd::Float64, tot_type1::Float64, pop_size::Float64)::Float64
   type2s = Int(type2s)
   type1s_reqd = AlterForIntegralChecks_Others(type1s_reqd)
   tot_type1 = AlterForIntegralChecks_Others(tot_type1)
   pop_size = AlterForIntegralChecks_Others(pop_size)
   if (type1s_reqd <= 0.0 || tot_type1 < type1s_reqd || tot_type1 > pop_size)
      cdf_neghypergeometric = [#VALUE!]
   elseif (tot_type1 + type2s > pop_size)
      cdf_neghypergeometric = 1.0
   else
      Dim ha1::Float64, hprob::Float64, hswap::Bool
      cdf_neghypergeometric = hypergeometric(type2s, type1s_reqd, pop_size - tot_type1 - type2s, tot_type1 - type1s_reqd, false, ha1, hprob, hswap)
   end
   cdf_neghypergeometric = GetRidOfMinusZeroes(cdf_neghypergeometric)
end

function comp_cdf_neghypergeometric(type2s::Float64, type1s_reqd::Float64, tot_type1::Float64, pop_size::Float64)::Float64
   type2s = Int(type2s)
   type1s_reqd = AlterForIntegralChecks_Others(type1s_reqd)
   tot_type1 = AlterForIntegralChecks_Others(tot_type1)
   pop_size = AlterForIntegralChecks_Others(pop_size)
   if (type1s_reqd <= 0.0 || tot_type1 < type1s_reqd || tot_type1 > pop_size)
      comp_cdf_neghypergeometric = [#VALUE!]
   elseif (tot_type1 + type2s > pop_size)
      comp_cdf_neghypergeometric = 0.0
   else
      Dim ha1::Float64, hprob::Float64, hswap::Bool
      comp_cdf_neghypergeometric = hypergeometric(type2s, type1s_reqd, pop_size - tot_type1 - type2s, tot_type1 - type1s_reqd, true, ha1, hprob, hswap)
   end
   comp_cdf_neghypergeometric = GetRidOfMinusZeroes(comp_cdf_neghypergeometric)
end

function crit_neghypergeometric(type1s_reqd::Float64, tot_type1::Float64, pop_size::Float64, crit_prob::Float64)::Float64
   type1s_reqd = AlterForIntegralChecks_Others(type1s_reqd)
   tot_type1 = AlterForIntegralChecks_Others(tot_type1)
   pop_size = AlterForIntegralChecks_Others(pop_size)
   if (crit_prob < 0.0 || crit_prob > 1.0)
      crit_neghypergeometric = [#VALUE!]
   elseif (type1s_reqd < 0.0 || tot_type1 < type1s_reqd || tot_type1 > pop_size)
      crit_neghypergeometric = [#VALUE!]
   elseif (crit_prob = 0.0)
      crit_neghypergeometric = [#VALUE!]
   elseif (pop_size = tot_type1)
      crit_neghypergeometric = 0.0
   elseif (crit_prob = 1.0)
      crit_neghypergeometric = pop_size - tot_type1
   else
      Dim ha1::Float64, hprob::Float64, hswap::Bool
      Dim i::Float64, pr::Float64
      crit_neghypergeometric = critneghyperg(type1s_reqd, tot_type1, pop_size, crit_prob)
      i = crit_neghypergeometric
      pr = hypergeometric(i, type1s_reqd, pop_size - tot_type1 - i, tot_type1 - type1s_reqd, false, ha1, hprob, hswap)
      if (pr = crit_prob)
      elseif (pr > crit_prob)
         i = i - 1.0
         pr = hypergeometric(i, type1s_reqd, pop_size - tot_type1 - i, tot_type1 - type1s_reqd, false, ha1, hprob, hswap)
         if (pr >= crit_prob)
            crit_neghypergeometric = i
         end
      else
         crit_neghypergeometric = i + 1.0
      end
   end
   crit_neghypergeometric = GetRidOfMinusZeroes(crit_neghypergeometric)
end

function comp_crit_neghypergeometric(type1s_reqd::Float64, tot_type1::Float64, pop_size::Float64, crit_prob::Float64)::Float64
   type1s_reqd = AlterForIntegralChecks_Others(type1s_reqd)
   tot_type1 = AlterForIntegralChecks_Others(tot_type1)
   pop_size = AlterForIntegralChecks_Others(pop_size)
   if (crit_prob < 0.0 || crit_prob > 1.0)
      comp_crit_neghypergeometric = [#VALUE!]
   elseif (type1s_reqd <= 0.0 || tot_type1 < type1s_reqd || tot_type1 > pop_size)
      comp_crit_neghypergeometric = [#VALUE!]
   elseif (crit_prob = 1.0)
      comp_crit_neghypergeometric = [#VALUE!]
   elseif (crit_prob = 0.0 || pop_size = tot_type1)
      comp_crit_neghypergeometric = pop_size - tot_type1
   else
      Dim ha1::Float64, hprob::Float64, hswap::Bool
      Dim i::Float64, pr::Float64
      comp_crit_neghypergeometric = critcompneghyperg(type1s_reqd, tot_type1, pop_size, crit_prob)
      i = comp_crit_neghypergeometric
      pr = hypergeometric(i, type1s_reqd, pop_size - tot_type1 - i, tot_type1 - type1s_reqd, true, ha1, hprob, hswap)
      if (pr = crit_prob)
      elseif (pr < crit_prob)
         i = i - 1.0
         pr = hypergeometric(i, type1s_reqd, pop_size - tot_type1 - i, tot_type1 - type1s_reqd, true, ha1, hprob, hswap)
         if (pr <= crit_prob)
            comp_crit_neghypergeometric = i
         end
      else
         comp_crit_neghypergeometric = i + 1.0
      end
   end
   comp_crit_neghypergeometric = GetRidOfMinusZeroes(comp_crit_neghypergeometric)
end

function lcb_neghypergeometric(type2s::Float64, type1s_reqd::Float64, pop_size::Float64, prob::Float64)::Float64
   type2s = AlterForIntegralChecks_Others(type2s)
   type1s_reqd = AlterForIntegralChecks_Others(type1s_reqd)
   pop_size = AlterForIntegralChecks_Others(pop_size)
   if (prob < 0.0 || prob > 1.0)
      lcb_neghypergeometric = [#VALUE!]
   elseif (type1s_reqd <= 0.0 || type1s_reqd > pop_size || type2s > pop_size - type1s_reqd)
      lcb_neghypergeometric = [#VALUE!]
   elseif (prob = 0.0 || pop_size = type2s + type1s_reqd)
      lcb_neghypergeometric = type1s_reqd
   elseif (prob = 1.0)
      lcb_neghypergeometric = pop_size - type2s
   elseif (prob < 0.5)
      lcb_neghypergeometric = critneghyperg(type1s_reqd, type2s + type1s_reqd, pop_size, prob * (1.000000000001)) + type1s_reqd
   else
      lcb_neghypergeometric = critcompneghyperg(type1s_reqd, type2s + type1s_reqd, pop_size, (1.0 - prob) * (1.0 - 0.000000000001)) + type1s_reqd
   end
   lcb_neghypergeometric = GetRidOfMinusZeroes(lcb_neghypergeometric)
end

function ucb_neghypergeometric(type2s::Float64, type1s_reqd::Float64, pop_size::Float64, prob::Float64)::Float64
   type2s = AlterForIntegralChecks_Others(type2s)
   type1s_reqd = AlterForIntegralChecks_Others(type1s_reqd)
   pop_size = AlterForIntegralChecks_Others(pop_size)
   if (prob < 0.0 || prob > 1.0)
      ucb_neghypergeometric = [#VALUE!]
   elseif (type1s_reqd <= 0.0 || type1s_reqd > pop_size || type2s > pop_size - type1s_reqd)
      ucb_neghypergeometric = [#VALUE!]
   elseif (prob = 0.0 || type2s = 0.0 || pop_size = type2s + type1s_reqd)
      ucb_neghypergeometric = pop_size - type2s
   elseif (prob = 1.0)
      ucb_neghypergeometric = type1s_reqd
   elseif (prob < 0.5)
      ucb_neghypergeometric = critcompneghyperg(type1s_reqd, type2s + type1s_reqd - 1.0, pop_size, prob * (1.0 - 0.000000000001)) + type1s_reqd - 1.0
   else
      ucb_neghypergeometric = critneghyperg(type1s_reqd, type2s + type1s_reqd - 1.0, pop_size, (1.0 - prob) * (1.000000000001)) + type1s_reqd - 1.0
   end
   ucb_neghypergeometric = GetRidOfMinusZeroes(ucb_neghypergeometric)
end

function pdf_triangular(x::Float64, Min::Float64, mode::Float64, Max::Float64)::Float64
   if (Min > mode || mode > Max)
      pdf_triangular = [#VALUE!]
   elseif (x <= Min || x >= Max)
      pdf_triangular = 0.0
   elseif (x <= mode)
      pdf_triangular = 2.0 * (x - Min) / (mode - Min) / (Max - Min)
   else
      pdf_triangular = 2.0 * (Max - x) / (Max - mode) / (Max - Min)
   end
end

function cdf_triangular(x::Float64, Min::Float64, mode::Float64, Max::Float64)::Float64
   if (Min > mode || mode > Max)
      cdf_triangular = [#VALUE!]
   elseif (x <= Min)
      cdf_triangular = 0.0
   elseif (x >= Max)
      cdf_triangular = 1.0
   elseif (x <= mode)
      cdf_triangular = ((x - Min) / (mode - Min)) * ((x - Min) / (Max - Min))
   else
      cdf_triangular = (mode - Min) / (Max - Min) + (1 + (Max - x) / (Max - mode)) * ((x - mode) / (Max - Min))
   end
end

function comp_cdf_triangular(x::Float64, Min::Float64, mode::Float64, Max::Float64)::Float64
   if (Min > mode || mode > Max)
      comp_cdf_triangular = [#VALUE!]
   elseif (x <= Min)
      comp_cdf_triangular = 1.0
   elseif (x >= Max)
      comp_cdf_triangular = 0.0
   elseif (x <= mode)
      comp_cdf_triangular = (Max - mode) / (Max - Min) + (1 + (x - Min) / (mode - Min)) * ((mode - x) / (Max - Min))
   else
      comp_cdf_triangular = ((Max - x) / (Max - mode)) * ((Max - x) / (Max - Min))
   end
end

function inv_triangular(prob::Float64, Min::Float64, mode::Float64, Max::Float64)::Float64
Dim temp::Float64
   if (prob < 0.0 || prob > 1.0 || Min > mode || mode > Max)
      inv_triangular = [#VALUE!]
   elseif (prob <= (mode - Min) / (Max - Min))
      inv_triangular = Min + abs2(prob) * abs2(mode - Min) * abs2(Max - Min)
   else
      if prob > 0.5
         inv_triangular = Max - abs2(1.0 - prob) * abs2(Max - Min) * abs2(Max - mode)
      else
         temp = (Max - mode) / (Max - Min)
         inv_triangular = mode + (Max - mode) * (prob - (mode - Min) / (Max - Min)) / (temp + abs2(temp * (1.0 - prob)))
      end
   end
end

function comp_inv_triangular(prob::Float64, Min::Float64, mode::Float64, Max::Float64)::Float64
Dim temp::Float64
   if (prob < 0.0 || prob > 1.0 || Min > mode || mode > Max)
      comp_inv_triangular = [#VALUE!]
   elseif (prob <= (Max - mode) / (Max - Min))
      comp_inv_triangular = Max - abs2(prob) * abs2(Max - mode) * abs2(Max - Min)
   else
      if prob > 0.5
         comp_inv_triangular = Min + abs2(1.0 - prob) * abs2(mode - Min) * abs2(Max - Min)
      else
         temp = (mode - Min) / (Max - Min)
         comp_inv_triangular = mode - (mode - Min) * (prob - (Max - mode) / (Max - Min)) / (temp + abs2(temp * (1.0 - prob)))
      end
   end
end

function pdf_exponential(x::Float64, lambda::Float64)::Float64
   if (lambda <= 0.0)
      pdf_exponential = [#VALUE!]
   elseif (x < 0.0)
      pdf_exponential = 0.0
   else
      pdf_exponential = exp(-lambda * x + log(lambda))
   end
   pdf_exponential = GetRidOfMinusZeroes(pdf_exponential)
end

function cdf_exponential(x::Float64, lambda::Float64)::Float64
   if (lambda <= 0.0)
      cdf_exponential = [#VALUE!]
   elseif (x < 0.0)
      cdf_exponential = 0.0
   else
      cdf_exponential = -expm1(-lambda * x)
   end
   cdf_exponential = GetRidOfMinusZeroes(cdf_exponential)
end

function comp_cdf_exponential(x::Float64, lambda::Float64)::Float64
   if (lambda <= 0.0)
      comp_cdf_exponential = [#VALUE!]
   elseif (x < 0.0)
      comp_cdf_exponential = 1.0
   else
      comp_cdf_exponential = exp(-lambda * x)
   end
   comp_cdf_exponential = GetRidOfMinusZeroes(comp_cdf_exponential)
end

function inv_exponential(prob::Float64, lambda::Float64)::Float64
   if (lambda <= 0.0 || prob < 0.0 || prob >= 1.0)
      inv_exponential = [#VALUE!]
   else
      inv_exponential = -log0(-prob) / lambda
   end
   inv_exponential = GetRidOfMinusZeroes(inv_exponential)
end

function comp_inv_exponential(prob::Float64, lambda::Float64)::Float64
   if (lambda <= 0.0 || prob <= 0.0 || prob > 1.0)
      comp_inv_exponential = [#VALUE!]
   else
      comp_inv_exponential = -log(prob) / lambda
   end
   comp_inv_exponential = GetRidOfMinusZeroes(comp_inv_exponential)
end

function pdf_normal(x::Float64)::Float64
   if (abs(x) < 40.0)
      pdf_normal = exp(-x * x * 0.5) * OneOverSqrTwoPi
   else
      pdf_normal = 0.0
   end
   pdf_normal = GetRidOfMinusZeroes(pdf_normal)
end

function cdf_normal(x::Float64)::Float64
   cdf_normal = cnormal(x)
   cdf_normal = GetRidOfMinusZeroes(cdf_normal)
end

function inv_normal(prob::Float64)::Float64
   if (prob <= 0.0 || prob >= 1.0)
      inv_normal = [#VALUE!]
   else
      inv_normal = invcnormal(prob)
   end
   inv_normal = GetRidOfMinusZeroes(inv_normal)
end

function pdf_chi_sq(x::Float64, df::Float64)::Float64
   df = AlterForIntegralChecks_df(df)
   pdf_chi_sq = pdf_gamma(x, df / 2.0, 2.0)
   pdf_chi_sq = GetRidOfMinusZeroes(pdf_chi_sq)
end

function cdf_chi_sq(x::Float64, df::Float64)::Float64
   df = AlterForIntegralChecks_df(df)
   if (df <= 0.0)
      cdf_chi_sq = [#VALUE!]
   elseif (x <= 0.0)
      cdf_chi_sq = 0.0
   else
      cdf_chi_sq = gamma(x / 2.0, df / 2.0)
   end
   cdf_chi_sq = GetRidOfMinusZeroes(cdf_chi_sq)
end

function comp_cdf_chi_sq(x::Float64, df::Float64)::Float64
   df = AlterForIntegralChecks_df(df)
   if (df <= 0.0)
      comp_cdf_chi_sq = [#VALUE!]
   elseif (x <= 0.0)
      comp_cdf_chi_sq = 1.0
   else
      comp_cdf_chi_sq = compgamma(x / 2.0, df / 2.0)
   end
   comp_cdf_chi_sq = GetRidOfMinusZeroes(comp_cdf_chi_sq)
end

function inv_chi_sq(prob::Float64, df::Float64)::Float64
   df = AlterForIntegralChecks_df(df)
   if (df <= 0.0 || prob < 0.0 || prob >= 1.0)
      inv_chi_sq = [#VALUE!]
   elseif (prob = 0.0)
      inv_chi_sq = 0.0
   else
      inv_chi_sq = 2.0 * invgamma(df / 2.0, prob)
   end
   inv_chi_sq = GetRidOfMinusZeroes(inv_chi_sq)
end

function comp_inv_chi_sq(prob::Float64, df::Float64)::Float64
   df = AlterForIntegralChecks_df(df)
   if (df <= 0.0 || prob <= 0.0 || prob > 1.0)
      comp_inv_chi_sq = [#VALUE!]
   elseif (prob = 1.0)
      comp_inv_chi_sq = 0.0
   else
      comp_inv_chi_sq = 2.0 * invcompgamma(df / 2.0, prob)
   end
   comp_inv_chi_sq = GetRidOfMinusZeroes(comp_inv_chi_sq)
end

function pdf_gamma(x::Float64, shape_param::Float64, scale_param::Float64)::Float64
   Dim xs::Float64
   if (shape_param <= 0.0 || scale_param <= 0.0)
      pdf_gamma = [#VALUE!]
   elseif (x < 0.0)
      pdf_gamma = 0.0
   elseif (x = 0.0)
      if (shape_param < 1.0)
         pdf_gamma = [#VALUE!]
      elseif (shape_param = 1.0)
         pdf_gamma = 1.0 / scale_param
      else
         pdf_gamma = 0.0
      end
   else
      xs = x / scale_param
      pdf_gamma = poissonTerm(shape_param, xs, xs - shape_param, log(shape_param) - log(x))
   end
   pdf_gamma = GetRidOfMinusZeroes(pdf_gamma)
end

function cdf_gamma(x::Float64, shape_param::Float64, scale_param::Float64)::Float64
   if (shape_param <= 0.0 || scale_param <= 0.0)
      cdf_gamma = [#VALUE!]
   elseif (x <= 0.0)
      cdf_gamma = 0.0
   else
      cdf_gamma = gamma(x / scale_param, shape_param)
   end
   cdf_gamma = GetRidOfMinusZeroes(cdf_gamma)
end

function comp_cdf_gamma(x::Float64, shape_param::Float64, scale_param::Float64)::Float64
   if (shape_param <= 0.0 || scale_param <= 0.0)
      comp_cdf_gamma = [#VALUE!]
   elseif (x <= 0.0)
      comp_cdf_gamma = 1.0
   else
      comp_cdf_gamma = compgamma(x / scale_param, shape_param)
   end
   comp_cdf_gamma = GetRidOfMinusZeroes(comp_cdf_gamma)
end

function inv_gamma(prob::Float64, shape_param::Float64, scale_param::Float64)::Float64
   if (shape_param <= 0.0 || scale_param <= 0.0 || prob < 0.0 || prob >= 1.0)
      inv_gamma = [#VALUE!]
   elseif (prob = 0.0)
      inv_gamma = 0.0
   else
      inv_gamma = scale_param * invgamma(shape_param, prob)
   end
   inv_gamma = GetRidOfMinusZeroes(inv_gamma)
end

function comp_inv_gamma(prob::Float64, shape_param::Float64, scale_param::Float64)::Float64
   if (shape_param <= 0.0 || scale_param <= 0.0 || prob <= 0.0 || prob > 1.0)
      comp_inv_gamma = [#VALUE!]
   elseif (prob = 1.0)
      comp_inv_gamma = 0.0
   else
      comp_inv_gamma = scale_param * invcompgamma(shape_param, prob)
   end
   comp_inv_gamma = GetRidOfMinusZeroes(comp_inv_gamma)
end

function  pdftdist(x::Float64, k::Float64)::Float64
#//Probability density for a variate from t-distribution with k degress of freedom
   Dim a::Float64, x2::Float64, k2::Float64, logterm::Float64, c5::Float64
   if (k <= 0.0)
      pdftdist = [#VALUE!]
   elseif (k > 1E+30)
      pdftdist = pdf_normal(x)
   else
      if abs(x) >= min(1.0, k)
         k2 = k / x
         x2 = x + k2
         k2 = k2 / x2
         x2 = x / x2
      else
         x2 = x * x
         k2 = k + x2
         x2 = x2 / k2
         k2 = k / k2
      end
      if (k2 < cSmall)
         logterm = log(k) - 2.0 * log(abs(x))
      elseif (abs(x2) < 0.5)
         logterm = log0(-x2)
      else
         logterm = log(k2)
      end
      a = k * 0.5
      c5 = -1.0 / (k + 2.0)
      pdftdist = exp((a + 0.5) * logterm + a * log1(c5) - c5 + lfbaccdif1(0.5, a - 0.5)) * abs2(a / ((1.0 + a))) * OneOverSqrTwoPi
   end
end

function pdf_tdist(x::Float64, df::Float64)::Float64
   df = AlterForIntegralChecks_df(df)
   pdf_tdist = pdftdist(x, df)
   pdf_tdist = GetRidOfMinusZeroes(pdf_tdist)
end

function cdf_tdist(x::Float64, df::Float64)::Float64
   Dim tdistDensity::Float64
   df = AlterForIntegralChecks_df(df)
   if (df <= 0.0)
      cdf_tdist = [#VALUE!]
   else
      cdf_tdist = tdist(x, df, tdistDensity)
   end
   cdf_tdist = GetRidOfMinusZeroes(cdf_tdist)
end

function inv_tdist(prob::Float64, df::Float64)::Float64
   df = AlterForIntegralChecks_df(df)
   if (df <= 0.0)
      inv_tdist = [#VALUE!]
   elseif (prob <= 0.0 || prob >= 1.0)
      inv_tdist = [#VALUE!]
   else
      inv_tdist = invtdist(prob, df)
   end
   inv_tdist = GetRidOfMinusZeroes(inv_tdist)
end

function pdf_fdist(x::Float64, df1::Float64, df2::Float64)::Float64
   df1 = AlterForIntegralChecks_df(df1)
   df2 = AlterForIntegralChecks_df(df2)
   if (df1 <= 0.0 || df2 <= 0.0)
      pdf_fdist = [#VALUE!]
   elseif (x < 0.0)
      pdf_fdist = 0.0
   elseif (x = 0.0 && df1 > 2.0)
      pdf_fdist = 0.0
   elseif (x = 0.0 && df1 < 2.0)
      pdf_fdist = [#VALUE!]
   elseif (x = 0.0)
      pdf_fdist = 1.0
   else
      Dim p::Float64, q::Float64
      if x > 1.0
         q = df2 / x
         p = q + df1
         q = q / p
         p = df1 / p
      else
         p = df1 * x
         q = df2 + p
         p = p / q
         q = df2 / q
      end
      #if p < cSmall && x <> 0.0 || q < cSmall
      #   pdf_fdist = [#VALUE!]
      #   Exit Function
      #end
      df2 = df2 / 2.0
      df1 = df1 / 2.0
      if (df1 >= 1.0)
         df1 = df1 - 1.0
         pdf_fdist = binomialTerm(df1, df2, p, q, df2 * p - df1 * q, log((df1 + 1.0) * q))
      else
         pdf_fdist = df1 * df1 * q / (p * (df1 + df2)) * binomialTerm(df1, df2, p, q, df2 * p - df1 * q, 0.0)
      end
   end
   pdf_fdist = GetRidOfMinusZeroes(pdf_fdist)
end

function cdf_fdist(x::Float64, df1::Float64, df2::Float64)::Float64
   df1 = AlterForIntegralChecks_df(df1)
   df2 = AlterForIntegralChecks_df(df2)
   if (df1 <= 0.0 || df2 <= 0.0)
      cdf_fdist = [#VALUE!]
   elseif (x <= 0.0)
      cdf_fdist = 0.0
   else
      Dim p::Float64, q::Float64
      if x > 1.0
         q = df2 / x
         p = q + df1
         q = q / p
         p = df1 / p
      else
         p = df1 * x
         q = df2 + p
         p = p / q
         q = df2 / q
      end
      #if p < cSmall && x <> 0.0 || q < cSmall
      #   cdf_fdist = [#VALUE!]
      #   Exit Function
      #end
      df2 = df2 / 2.0
      df1 = df1 / 2.0
      if (p < 0.5)
          cdf_fdist = beta(p, df1, df2)
      else
          cdf_fdist = compbeta(q, df2, df1)
      end
   end
   cdf_fdist = GetRidOfMinusZeroes(cdf_fdist)
end

function comp_cdf_fdist(x::Float64, df1::Float64, df2::Float64)::Float64
   df1 = AlterForIntegralChecks_df(df1)
   df2 = AlterForIntegralChecks_df(df2)
   if (df1 <= 0.0 || df2 <= 0.0)
      comp_cdf_fdist = [#VALUE!]
   elseif (x <= 0.0)
      comp_cdf_fdist = 1.0
   else
      Dim p::Float64, q::Float64
      if x > 1.0
         q = df2 / x
         p = q + df1
         q = q / p
         p = df1 / p
      else
         p = df1 * x
         q = df2 + p
         p = p / q
         q = df2 / q
      end
      #if p < cSmall && x <> 0.0 || q < cSmall
      #   comp_cdf_fdist = [#VALUE!]
      #   Exit Function
      #end
      df2 = df2 / 2.0
      df1 = df1 / 2.0
      if (p < 0.5)
          comp_cdf_fdist = compbeta(p, df1, df2)
      else
          comp_cdf_fdist = beta(q, df2, df1)
      end
   end
   comp_cdf_fdist = GetRidOfMinusZeroes(comp_cdf_fdist)
end

function inv_fdist(prob::Float64, df1::Float64, df2::Float64)::Float64
   df1 = AlterForIntegralChecks_df(df1)
   df2 = AlterForIntegralChecks_df(df2)
   if (df1 <= 0.0 || df2 <= 0.0 || prob < 0.0 || prob >= 1.0)
      inv_fdist = [#VALUE!]
   elseif (prob = 0.0)
      inv_fdist = 0.0
   else
      Dim temp::Float64, oneMinusP::Float64
      df1 = df1 / 2.0
      df2 = df2 / 2.0
      temp = invbeta(df1, df2, prob, oneMinusP)
      inv_fdist = df2 * temp / (df1 * oneMinusP)
      #if oneMinusP < cSmall inv_fdist = [#VALUE!]
   end
   inv_fdist = GetRidOfMinusZeroes(inv_fdist)
end

function comp_inv_fdist(prob::Float64, df1::Float64, df2::Float64)::Float64
   df1 = AlterForIntegralChecks_df(df1)
   df2 = AlterForIntegralChecks_df(df2)
   if (df1 <= 0.0 || df2 <= 0.0 || prob <= 0.0 || prob > 1.0)
      comp_inv_fdist = [#VALUE!]
   elseif (prob = 1.0)
      comp_inv_fdist = 0.0
   else
      Dim temp::Float64, oneMinusP::Float64
      df1 = df1 / 2.0
      df2 = df2 / 2.0
      temp = invcompbeta(df1, df2, prob, oneMinusP)
      comp_inv_fdist = df2 * temp / (df1 * oneMinusP)
      #if oneMinusP < cSmall comp_inv_fdist = [#VALUE!]
   end
   comp_inv_fdist = GetRidOfMinusZeroes(comp_inv_fdist)
end

function pdf_beta(x::Float64, shape_param1::Float64, shape_param2::Float64)::Float64
   if (shape_param1 <= 0.0 || shape_param2 <= 0.0)
      pdf_beta = [#VALUE!]
   elseif (x < 0.0 || x > 1.0)
      pdf_beta = 0.0
   elseif (x = 0.0 && shape_param1 < 1.0 || x = 1.0 && shape_param2 < 1.0)
      pdf_beta = [#VALUE!]
   elseif (x = 0.0 && shape_param1 = 1.0)
      pdf_beta = shape_param2
   elseif (x = 1.0 && shape_param2 = 1.0)
      pdf_beta = shape_param1
   elseif ((x = 0.0) || (x = 1.0))
      pdf_beta = 0.0
   else
      Dim mx::Float64, mn::Float64
      mx = max(shape_param1, shape_param2)
      mn = min(shape_param1, shape_param2)
      pdf_beta = (binomialTerm(shape_param1, shape_param2, x, 1.0 - x, (shape_param1 * (x - 1.0) + x * shape_param2), 0.0) * mx / (mn + mx)) * mn / (x * (1.0 - x))
   end
   pdf_beta = GetRidOfMinusZeroes(pdf_beta)
end

function cdf_beta(x::Float64, shape_param1::Float64, shape_param2::Float64)::Float64
   if (shape_param1 <= 0.0 || shape_param2 <= 0.0)
      cdf_beta = [#VALUE!]
   elseif (x <= 0.0)
      cdf_beta = 0.0
   elseif (x >= 1.0)
      cdf_beta = 1.0
   else
      cdf_beta = beta(x, shape_param1, shape_param2)
   end
   cdf_beta = GetRidOfMinusZeroes(cdf_beta)
end

function comp_cdf_beta(x::Float64, shape_param1::Float64, shape_param2::Float64)::Float64
   if (shape_param1 <= 0.0 || shape_param2 <= 0.0)
      comp_cdf_beta = [#VALUE!]
   elseif (x <= 0.0)
      comp_cdf_beta = 1.0
   elseif (x >= 1.0)
      comp_cdf_beta = 0.0
   else
      comp_cdf_beta = compbeta(x, shape_param1, shape_param2)
   end
   comp_cdf_beta = GetRidOfMinusZeroes(comp_cdf_beta)
end

function inv_beta(prob::Float64, shape_param1::Float64, shape_param2::Float64)::Float64
   if (shape_param1 <= 0.0 || shape_param2 <= 0.0 || prob < 0.0 || prob > 1.0)
      inv_beta = [#VALUE!]
   else
      Dim oneMinusP::Float64
      inv_beta = invbeta(shape_param1, shape_param2, prob, oneMinusP)
   end
   inv_beta = GetRidOfMinusZeroes(inv_beta)
end

function comp_inv_beta(prob::Float64, shape_param1::Float64, shape_param2::Float64)::Float64
   if (shape_param1 <= 0.0 || shape_param2 <= 0.0 || prob < 0.0 || prob > 1.0)
      comp_inv_beta = [#VALUE!]
   else
      Dim oneMinusP::Float64
      comp_inv_beta = invcompbeta(shape_param1, shape_param2, prob, oneMinusP)
   end
   comp_inv_beta = GetRidOfMinusZeroes(comp_inv_beta)
end

function  gamma_nc1(x::Float64, a::Float64, nc::Float64, ByRef nc_derivative::Float64)::Float64
   Dim aa::Float64, bb::Float64, nc_dtemp::Float64
   Dim n::Float64, p::Float64, w::Float64, s::Float64, ps::Float64
   Dim result::Float64, term::Float64, ptx::Float64, ptnc::Float64
   if a <= 1.0 && x <= 1.0
      n = a + abs2(a ^ 2 + 4.0 * nc * x)
      if n > 0.0 n = Int(2.0 * nc * x / n)
   elseif a > x
      n = x / a
      n = Int(2.0 * nc * n / (1.0 + abs2(1.0 + 4.0 * n * (nc / a))))
   elseif x >= a
      n = a / x
      n = Int(2.0 * nc / (n + abs2(n ^ 2 + 4.0 * (nc / x))))
   else
      Debug.Print x, a, nc
   end
   aa = n + a
   bb = n
   ptnc = poissonTerm(n, nc, nc - n, 0.0)
   ptx = poissonTerm(aa, x, x - aa, 0.0)
   aa = aa + 1.0
   bb = bb + 1.0
   p = nc / bb
   ps = p
   nc_derivative = ps
   s = x / aa
   w = p
   term = s * w
   result = term
   if ptx > 0.0
     while (((term > 0.000000000000001 * result) && (p > 1E-16 * w)) || (ps > 1E-16 * nc_derivative))
       aa = aa + 1.0
       bb = bb + 1.0
       p = nc / bb * p
       ps = p * s
       nc_derivative = nc_derivative + ps
       s = x / aa * s
       w = w + p
       term = s * w
       result = result + term
     end
     w = w * ptnc
   else
     w = comppoisson(n, nc, nc - n)
   end
   gamma_nc1 = result * ptx * ptnc + comppoisson(a + bb, x, (x - a) - bb) * w
   ps = 1.0
   nc_dtemp = 0.0
   aa = n + a
   bb = n
   p = 1.0
   s = ptx
   w = gamma(x, aa)
   term = p * w
   result = term
   while bb > 0.0 && ((term > 0.000000000000001 * result) || (ps > 1E-16 * nc_dtemp))
       s = aa / x * s
       ps = p * s
       nc_dtemp = nc_dtemp + ps
       p = bb / nc * p
       w = w + s
       term = p * w
       result = result + term
       aa = aa - 1.0
       bb = bb - 1.0
   end
   if bb = 0.0 aa = a
   if n > 0.0
      nc_dtemp = nc_derivative * ptx + nc_dtemp + p * aa / x * s
   else
      nc_dtemp = poissonTerm(aa, x, x - aa, log(nc_derivative * x + aa) - log(x))
   end
   gamma_nc1 = gamma_nc1 + result * ptnc + cpoisson(bb - 1.0, nc, nc - bb + 1.0) * w
   if nc_dtemp = 0.0
      nc_derivative = 0.0
   else
      nc_derivative = poissonTerm(n, nc, nc - n, log(nc_dtemp))
   end
end

function  comp_gamma_nc1(x::Float64, a::Float64, nc::Float64, ByRef nc_derivative::Float64)::Float64
   Dim aa::Float64, bb::Float64, nc_dtemp::Float64
   Dim n::Float64, p::Float64, w::Float64, s::Float64, ps::Float64
   Dim result::Float64, term::Float64, ptx::Float64, ptnc::Float64
   if a <= 1.0 && x <= 1.0
      n = a + abs2(a ^ 2 + 4.0 * nc * x)
      if n > 0.0 n = Int(2.0 * nc * x / n)
   elseif a > x
      n = x / a
      n = Int(2.0 * nc * n / (1.0 + abs2(1.0 + 4.0 * n * (nc / a))))
   elseif x >= a
      n = a / x
      n = Int(2.0 * nc / (n + abs2(n ^ 2 + 4.0 * (nc / x))))
   else
      Debug.Print x, a, nc
   end
   aa = n + a
   bb = n
   ptnc = poissonTerm(n, nc, nc - n, 0.0)
   ptx = poissonTerm(aa, x, x - aa, 0.0)
   s = 1.0
   ps = 1.0
   nc_dtemp = 0.0
   p = 1.0
   w = p
   term = 1.0
   result = 0.0
   if ptx > 0.0
     while bb > 0.0 && (((term > 0.000000000000001 * result) && (p > 1E-16 * w)) || (ps > 1E-16 * nc_dtemp))
      s = aa / x * s
      ps = p * s
      nc_dtemp = nc_dtemp + ps
      p = bb / nc * p
      term = s * w
      result = result + term
      w = w + p
      aa = aa - 1.0
      bb = bb - 1.0
     end
     w = w * ptnc
   else
     w = cpoisson(n, nc, nc - n)
   end
   if bb = 0.0 aa = a
   if n > 0.0
      nc_dtemp = (nc_dtemp + p * aa / x * s) * ptx
   elseif aa = 0 && x > 0
      nc_dtemp = 0.0
   else
      nc_dtemp = poissonTerm(aa, x, x - aa, log(aa) - log(x))
   end
   comp_gamma_nc1 = result * ptx * ptnc + compgamma(x, aa) * w
   aa = n + a
   bb = n
   ps = 1.0
   nc_derivative = 0.0
   p = 1.0
   s = ptx
   w = compgamma(x, aa)
   term = 0.0
   result = term
   Do
       w = w + s
       aa = aa + 1.0
       bb = bb + 1.0
       p = nc / bb * p
       ps = p * s
       nc_derivative = nc_derivative + ps
       s = x / aa * s
       term = p * w
       result = result + term
   end while (((term > 0.000000000000001 * result) && (s > 1E-16 * w)) || (ps > 1E-16 * nc_derivative))
   comp_gamma_nc1 = comp_gamma_nc1 + result * ptnc + comppoisson(bb, nc, nc - bb) * w
   nc_dtemp = nc_derivative + nc_dtemp
   if nc_dtemp = 0.0
      nc_derivative = 0.0
   else
      nc_derivative = poissonTerm(n, nc, nc - n, log(nc_dtemp))
   end
end

function  inv_gamma_nc1(prob::Float64, a::Float64, nc::Float64)::Float64
#Uses approx in A&S 26.4.27 for to get initial estimate the modified NR to improve it.
Dim x::Float64, pr::Float64, dif::Float64
Dim hi::Float64, lo::Float64, nc_derivative::Float64
   if (prob > 0.5)
      inv_gamma_nc1 = comp_inv_gamma_nc1(1.0 - prob, a, nc)
      Exit Function
   end

   lo = 0.0
   hi = 1E+308
   pr = exp(-nc)
   if pr > prob
      if 2.0 * prob > pr
         x = comp_inv_gamma((pr - prob) / pr, a + cSmall, 1.0)
      else
         x = inv_gamma(prob / pr, a + cSmall, 1.0)
      end
      if x < cSmall
         x = cSmall
         pr = gamma_nc1(x, a, nc, nc_derivative)
         if pr > prob
            inv_gamma_nc1 = 0.0
            Exit Function
         end
      end
   else
      x = inv_gamma(prob, (a + nc) / (1.0 + nc / (a + nc)), 1.0)
      x = x * (1.0 + nc / (a + nc))
   end
   dif = x
   Do
      pr = gamma_nc1(x, a, nc, nc_derivative)
      if pr < 3E-308 && nc_derivative = 0.0
         lo = x
         dif = dif / 2.0
         x = x - dif
      elseif nc_derivative = 0.0
         hi = x
         dif = dif / 2.0
         x = x - dif
      else
         if pr < prob
            lo = x
         else
            hi = x
         end
         dif = -(pr / nc_derivative) * logdif(pr, prob)
         if x + dif < lo
            dif = (lo - x) / 2.0
         elseif x + dif > hi
            dif = (hi - x) / 2.0
         end
         x = x + dif
      end
   end while ((abs(pr - prob) > prob * 0.00000000000001) && (abs(dif) > abs(x) * 0.0000000001))
   inv_gamma_nc1 = x
end

function  comp_inv_gamma_nc1(prob::Float64, a::Float64, nc::Float64)::Float64
#Uses approx in A&S 26.4.27 for to get initial estimate the modified NR to improve it.
Dim x::Float64, pr::Float64, dif::Float64
Dim hi::Float64, lo::Float64, nc_derivative::Float64
   if (prob > 0.5)
      comp_inv_gamma_nc1 = inv_gamma_nc1(1.0 - prob, a, nc)
      Exit Function
   end

   lo = 0.0
   hi = 1E+308
   pr = exp(-nc)
   if pr > prob
      x = comp_inv_gamma(prob / pr, a + cSmall, 1.0) # Is this as small as x could be?
   else
      x = comp_inv_gamma(prob, (a + nc) / (1.0 + nc / (a + nc)), 1.0)
      x = x * (1.0 + nc / (a + nc))
   end
   if x < cSmall x = cSmall
   dif = x
   Do
      pr = comp_gamma_nc1(x, a, nc, nc_derivative)
      if pr < 3E-308 && nc_derivative = 0.0
         hi = x
         dif = dif / 2.0
         x = x - dif
      elseif nc_derivative = 0.0
         lo = x
         dif = dif / 2.0
         x = x - dif
      else
         if pr < prob
            hi = x
         else
            lo = x
         end
         dif = (pr / nc_derivative) * logdif(pr, prob)
         if x + dif < lo
            dif = (lo - x) / 2.0
         elseif x + dif > hi
            dif = (hi - x) / 2.0
         end
         x = x + dif
      end
   end while ((abs(pr - prob) > prob * 0.00000000000001) && (abs(dif) > abs(x) * 0.0000000001))
   comp_inv_gamma_nc1 = x
end

function  ncp_gamma_nc1(prob::Float64, x::Float64, a::Float64)::Float64
#Uses Normal approx for difference of 2 poisson distributed variables  to get initial estimate the modified NR to improve it.
Dim ncp::Float64, pr::Float64, dif::Float64, temp::Float64, deriv::Float64, b::Float64, sqarg::Float64, checked_nc_limit::Bool, checked_0_limit::Bool
Dim hi::Float64, lo::Float64
   if (prob > 0.5)
      ncp_gamma_nc1 = comp_ncp_gamma_nc1(1.0 - prob, x, a)
      Exit Function
   end

   lo = 0.0
   hi = nc_limit
   checked_0_limit = false
   checked_nc_limit = false
   temp = inv_normal(prob) ^ 2
   b = 2.0 * (x - a) + temp
   sqarg = b ^ 2 - 4 * ((x - a) ^ 2 - temp * x)
   if sqarg < 0
      ncp = b / 2
   else
      ncp = (b + abs2(sqarg)) / 2
   end
   ncp = max(0.0, min(ncp, nc_limit))
   if ncp = 0.0
      pr = cdf_gamma_nc(x, a, 0.0)
      if pr < prob
         if (inv_gamma(prob, a, 1) <= x)
            ncp_gamma_nc1 = 0.0
         else
            ncp_gamma_nc1 = [#VALUE!]
         end
         Exit Function
      else
         checked_0_limit = true
      end
   elseif ncp = nc_limit
      pr = cdf_gamma_nc(x, a, ncp)
      if pr > prob
         ncp_gamma_nc1 = [#VALUE!]
         Exit Function
      else
         checked_nc_limit = true
      end
   end
   dif = ncp
   Do
      pr = cdf_gamma_nc(x, a, ncp)
      #Debug.Print ncp, pr, prob
      deriv = pdf_gamma_nc(x, a + 1.0, ncp)
      if pr < 3E-308 && deriv = 0.0
         hi = ncp
         dif = dif / 2.0
         ncp = ncp - dif
      elseif deriv = 0.0
         lo = ncp
         dif = dif / 2.0
         ncp = ncp - dif
      else
         if pr < prob
            hi = ncp
         else
            lo = ncp
         end
         dif = (pr / deriv) * logdif(pr, prob)
         if ncp + dif < lo
            dif = (lo - ncp) / 2.0
            if Not checked_0_limit && (lo = 0.0)
               temp = cdf_gamma_nc(x, a, lo)
               if temp < prob
                  if (inv_gamma(prob, a, 1) <= x)
                     ncp_gamma_nc1 = 0.0
                  else
                     ncp_gamma_nc1 = [#VALUE!]
                  end
                  Exit Function
               else
                  checked_0_limit = true
               end
            end
         elseif ncp + dif > hi
            dif = (hi - ncp) / 2.0
            if Not checked_nc_limit && (hi = nc_limit)
               pr = cdf_gamma_nc(x, a, hi)
               if pr > prob
                  ncp_gamma_nc1 = [#VALUE!]
                  Exit Function
               else
                  ncp = hi
                  deriv = pdf_gamma_nc(x, a + 1.0, ncp)
                  dif = (pr / deriv) * logdif(pr, prob)
                  if ncp + dif < lo
                     dif = (lo - ncp) / 2.0
                  end
                  checked_nc_limit = true
               end
            end
         end
         ncp = ncp + dif
      end
   end while ((abs(pr - prob) > prob * 0.00000000000001) && (abs(dif) > abs(ncp) * 0.0000000001))
   ncp_gamma_nc1 = ncp
   #Debug.Print "ncp_gamma_nc1", ncp_gamma_nc1
end

function  comp_ncp_gamma_nc1(prob::Float64, x::Float64, a::Float64)::Float64
#Uses Normal approx for difference of 2 poisson distributed variables  to get initial estimate the modified NR to improve it.
Dim ncp::Float64, pr::Float64, dif::Float64, temp::Float64, deriv::Float64, b::Float64, sqarg::Float64, checked_nc_limit::Bool, checked_0_limit::Bool
Dim hi::Float64, lo::Float64
   if (prob > 0.5)
      comp_ncp_gamma_nc1 = ncp_gamma_nc1(1.0 - prob, x, a)
      Exit Function
   end

   lo = 0.0
   hi = nc_limit
   checked_0_limit = false
   checked_nc_limit = false
   temp = inv_normal(prob) ^ 2
   b = 2.0 * (x - a) + temp
   sqarg = b ^ 2 - 4 * ((x - a) ^ 2 - temp * x)
   if sqarg < 0
      ncp = b / 2
   else
      ncp = (b - abs2(sqarg)) / 2
   end
   ncp = max(0.0, ncp)
   if ncp <= 1.0
      pr = comp_cdf_gamma_nc(x, a, 0.0)
      if pr > prob
         if (comp_inv_gamma(prob, a, 1) <= x)
            comp_ncp_gamma_nc1 = 0.0
         else
            comp_ncp_gamma_nc1 = [#VALUE!]
         end
         Exit Function
      else
         checked_0_limit = true
      end
      deriv = pdf_gamma_nc(x, a + 1.0, ncp)
      if deriv = 0.0
         ncp = nc_limit
      elseif a < 1
         ncp = (prob - pr) / deriv
         if ncp >= nc_limit
            ncp = -(pr / deriv) * logdif(pr, prob)
         end
      else
         ncp = -(pr / deriv) * logdif(pr, prob)
      end
   end
   ncp = min(ncp, nc_limit)
   if ncp = nc_limit
      pr = comp_cdf_gamma_nc(x, a, ncp)
      if pr < prob
         comp_ncp_gamma_nc1 = [#VALUE!]
         Exit Function
      else
         deriv = pdf_gamma_nc(x, a + 1.0, ncp)
         dif = -(pr / deriv) * logdif(pr, prob)
         if ncp + dif < lo
            dif = (lo - ncp) / 2.0
         end
         checked_nc_limit = true
      end
   end
   dif = ncp
   Do
      pr = comp_cdf_gamma_nc(x, a, ncp)
      #Debug.Print ncp, pr, prob
      deriv = pdf_gamma_nc(x, a + 1.0, ncp)
      if pr < 3E-308 && deriv = 0.0
         lo = ncp
         dif = dif / 2.0
         ncp = ncp + dif
      elseif deriv = 0.0
         hi = ncp
         dif = dif / 2.0
         ncp = ncp - dif
      else
         if pr < prob
            lo = ncp
         else
            hi = ncp
         end
         dif = -(pr / deriv) * logdif(pr, prob)
         if ncp + dif < lo
            dif = (lo - ncp) / 2.0
            if Not checked_0_limit && (lo = 0.0)
               temp = comp_cdf_gamma_nc(x, a, lo)
               if temp > prob
                  if (comp_inv_gamma(prob, a, 1) <= x)
                     comp_ncp_gamma_nc1 = 0.0
                  else
                     comp_ncp_gamma_nc1 = [#VALUE!]
                  end
                  Exit Function
               else
                  checked_0_limit = true
               end
            end
         elseif ncp + dif > hi
            if Not checked_nc_limit && (hi = nc_limit)
               ncp = hi
               pr = comp_cdf_gamma_nc(x, a, ncp)
               if pr < prob
                  comp_ncp_gamma_nc1 = [#VALUE!]
                  Exit Function
               else
                  deriv = pdf_gamma_nc(x, a + 1.0, ncp)
                  dif = -(pr / deriv) * logdif(pr, prob)
                  if ncp + dif < lo
                     dif = (lo - ncp) / 2.0
                  end
                  checked_nc_limit = true
               end
            else
               dif = (hi - ncp) / 2.0
            end
         end
         ncp = ncp + dif
      end
   end while ((abs(pr - prob) > prob * 0.00000000000001) && (abs(dif) > abs(ncp) * 0.0000000001))
   comp_ncp_gamma_nc1 = ncp
   #Debug.Print "comp_ncp_gamma_nc1", comp_ncp_gamma_nc1
end

function pdf_gamma_nc(x::Float64, shape_param::Float64, nc_param::Float64)::Float64
#// Calculate pdf of noncentral gamma
  Dim nc_derivative::Float64
  if (shape_param < 0.0) || (nc_param < 0.0) || (nc_param > nc_limit)
     pdf_gamma_nc = [#VALUE!]
  elseif (x < 0.0)
     pdf_gamma_nc = 0.0
  elseif (shape_param = 0.0 && nc_param = 0.0 && x > 0.0)
     pdf_gamma_nc = 0.0
  elseif (x = 0.0 || nc_param = 0.0)
     pdf_gamma_nc = exp(-nc_param) * pdf_gamma(x, shape_param, 1.0)
  elseif shape_param >= 1.0
     if x >= nc_param
        if (x < 1.0 || x <= shape_param + nc_param)
           pdf_gamma_nc = gamma_nc1(x, shape_param, nc_param, nc_derivative)
        else
           pdf_gamma_nc = comp_gamma_nc1(x, shape_param, nc_param, nc_derivative)
        end
        pdf_gamma_nc = nc_derivative
     else
        if (nc_param < 1.0 || nc_param <= shape_param + x)
           pdf_gamma_nc = gamma_nc1(nc_param, shape_param, x, nc_derivative)
        else
           pdf_gamma_nc = comp_gamma_nc1(nc_param, shape_param, x, nc_derivative)
        end
        if nc_derivative = 0.0
           pdf_gamma_nc = 0.0
        else
           pdf_gamma_nc = exp(log(nc_derivative) + (shape_param - 1.0) * (log(x) - log(nc_param)))
        end
     end
  else
     if x < nc_param
        if (x < 1.0 || x <= shape_param + nc_param)
           pdf_gamma_nc = gamma_nc1(x, shape_param, nc_param, nc_derivative)
        else
           pdf_gamma_nc = comp_gamma_nc1(x, shape_param, nc_param, nc_derivative)
        end
        pdf_gamma_nc = nc_derivative
     else
        if (nc_param < 1.0 || nc_param <= shape_param + x)
           pdf_gamma_nc = gamma_nc1(nc_param, shape_param, x, nc_derivative)
        else
           pdf_gamma_nc = comp_gamma_nc1(nc_param, shape_param, x, nc_derivative)
        end
        if nc_derivative = 0.0
           pdf_gamma_nc = 0.0
        else
           pdf_gamma_nc = exp(log(nc_derivative) + (shape_param - 1.0) * (log(x) - log(nc_param)))
        end
     end
  end
  pdf_gamma_nc = GetRidOfMinusZeroes(pdf_gamma_nc)
end

function cdf_gamma_nc(x::Float64, shape_param::Float64, nc_param::Float64)::Float64
#// Calculate cdf of noncentral gamma
  Dim nc_derivative::Float64
  if (shape_param < 0.0) || (nc_param < 0.0) || (nc_param > nc_limit)
     cdf_gamma_nc = [#VALUE!]
  elseif (x < 0.0)
     cdf_gamma_nc = 0.0
  elseif (x = 0.0 && shape_param = 0.0)
     cdf_gamma_nc = exp(-nc_param)
  elseif (shape_param + nc_param = 0.0)    # limit as shape_param+nc_param->0 is degenerate point mass at zero
     cdf_gamma_nc = 1.0                         # if fix central gamma, then works for degenerate poisson
  elseif (x = 0.0)
     cdf_gamma_nc = 0.0
  elseif (nc_param = 0.0)
     cdf_gamma_nc = gamma(x, shape_param)
  #elseif (shape_param = 0.0)              # extends Ruben (1974) and Cohen (1988) recurrence
  #   cdf_gamma_nc = ((x + shape_param + 2.0) * gamma_nc1(x, shape_param + 2.0, nc_param) + (nc_param - shape_param - 2.0) * gamma_nc1(x, shape_param + 4.0, nc_param) - nc_param * gamma_nc1(x, shape_param + 6.0, nc_param)) / x
  elseif (x < 1.0 || x <= shape_param + nc_param)
     cdf_gamma_nc = gamma_nc1(x, shape_param, nc_param, nc_derivative)
  else
     cdf_gamma_nc = 1.0 - comp_gamma_nc1(x, shape_param, nc_param, nc_derivative)
  end
  cdf_gamma_nc = GetRidOfMinusZeroes(cdf_gamma_nc)
end

function comp_cdf_gamma_nc(x::Float64, shape_param::Float64, nc_param::Float64)::Float64
#// Calculate 1-cdf of noncentral gamma
  Dim nc_derivative::Float64
  if (shape_param < 0.0) || (nc_param < 0.0) || (nc_param > nc_limit)
     comp_cdf_gamma_nc = [#VALUE!]
  elseif (x < 0.0)
     comp_cdf_gamma_nc = 1.0
  elseif (x = 0.0 && shape_param = 0.0)
     comp_cdf_gamma_nc = -expm1(-nc_param)
  elseif (shape_param + nc_param = 0.0)     # limit as shape_param+nc_param->0 is degenerate point mass at zero
     comp_cdf_gamma_nc = 0.0                     # if fix central gamma, then works for degenerate poisson
  elseif (x = 0.0)
     comp_cdf_gamma_nc = 1
  elseif (nc_param = 0.0)
     comp_cdf_gamma_nc = compgamma(x, shape_param)
  #elseif (shape_param = 0.0)              # extends Ruben (1974) and Cohen (1988) recurrence
  #   comp_cdf_gamma_nc = ((x + shape_param + 2.0) * comp_gamma_nc1(x, shape_param + 2.0, nc_param) + (nc_param - shape_param - 2.0) * comp_gamma_nc1(x, shape_param + 4.0, nc_param) - nc_param * comp_gamma_nc1(x, shape_param + 6.0, nc_param)) / x
  elseif (x < 1.0 || x >= shape_param + nc_param)
     comp_cdf_gamma_nc = comp_gamma_nc1(x, shape_param, nc_param, nc_derivative)
  else
     comp_cdf_gamma_nc = 1.0 - gamma_nc1(x, shape_param, nc_param, nc_derivative)
  end
  comp_cdf_gamma_nc = GetRidOfMinusZeroes(comp_cdf_gamma_nc)
end

function inv_gamma_nc(prob::Float64, shape_param::Float64, nc_param::Float64)::Float64
   if (shape_param < 0.0 || nc_param < 0.0 || nc_param > nc_limit || prob < 0.0 || prob >= 1.0)
      inv_gamma_nc = [#VALUE!]
   elseif (prob = 0.0 || shape_param = 0.0 && prob <= exp(-nc_param))
      inv_gamma_nc = 0.0
   else
      inv_gamma_nc = inv_gamma_nc1(prob, shape_param, nc_param)
   end
   inv_gamma_nc = GetRidOfMinusZeroes(inv_gamma_nc)
end

function comp_inv_gamma_nc(prob::Float64, shape_param::Float64, nc_param::Float64)::Float64
   if (shape_param < 0.0 || nc_param < 0.0 || nc_param > nc_limit || prob <= 0.0 || prob > 1.0)
      comp_inv_gamma_nc = [#VALUE!]
   elseif (prob = 1.0 || shape_param = 0.0 && prob >= -expm1(-nc_param))
      comp_inv_gamma_nc = 0.0
   else
      comp_inv_gamma_nc = comp_inv_gamma_nc1(prob, shape_param, nc_param)
   end
   comp_inv_gamma_nc = GetRidOfMinusZeroes(comp_inv_gamma_nc)
end

function ncp_gamma_nc(prob::Float64, x::Float64, shape_param::Float64)::Float64
   if (shape_param < 0.0 || x < 0.0 || prob <= 0.0 || prob > 1.0)
      ncp_gamma_nc = [#VALUE!]
   elseif (x = 0.0 && shape_param = 0.0)
      ncp_gamma_nc = -log(prob)
   elseif (shape_param = 0.0 && prob = 1.0)
      ncp_gamma_nc = 0.0
   elseif (x = 0.0 || prob = 1.0)
      ncp_gamma_nc = [#VALUE!]
   else
      ncp_gamma_nc = ncp_gamma_nc1(prob, x, shape_param)
   end
   ncp_gamma_nc = GetRidOfMinusZeroes(ncp_gamma_nc)
end

function comp_ncp_gamma_nc(prob::Float64, x::Float64, shape_param::Float64)::Float64
   if (shape_param < 0.0 || x < 0.0 || prob < 0.0 || prob >= 1.0)
      comp_ncp_gamma_nc = [#VALUE!]
   elseif (x = 0.0 && shape_param = 0.0)
      comp_ncp_gamma_nc = -log0(-prob)
   elseif (shape_param = 0.0 && prob = 0.0)
      comp_ncp_gamma_nc = 0.0
   elseif (x = 0.0 || prob = 0.0)
      comp_ncp_gamma_nc = [#VALUE!]
   else
      comp_ncp_gamma_nc = comp_ncp_gamma_nc1(prob, x, shape_param)
   end
   comp_ncp_gamma_nc = GetRidOfMinusZeroes(comp_ncp_gamma_nc)
end

function pdf_Chi2_nc(x::Float64, df::Float64, nc::Float64)::Float64
#// Calculate pdf of noncentral chi-square
  df = AlterForIntegralChecks_df(df)
  pdf_Chi2_nc = 0.5 * pdf_gamma_nc(x / 2.0, df / 2.0, nc / 2.0)
  pdf_Chi2_nc = GetRidOfMinusZeroes(pdf_Chi2_nc)
end

function cdf_Chi2_nc(x::Float64, df::Float64, nc::Float64)::Float64
#// Calculate cdf of noncentral chi-square
#//   parametrized per Johnson & Kotz, SAS, etc. so that cdf_Chi2_nc(x,df,nc) = cdf_gamma_nc(x/2,df/2,nc/2)
#//   if Xi ~ N(Di,1) independent, then sum(Xi,i=1..n) ~ Chi2_nc(n,nc) with nc=sum(Di,i=1..n)
#//   Note that Knusel, Graybill, etc. use a different noncentrality parameter lambda=nc/2
  df = AlterForIntegralChecks_df(df)
  cdf_Chi2_nc = cdf_gamma_nc(x / 2.0, df / 2.0, nc / 2.0)
  cdf_Chi2_nc = GetRidOfMinusZeroes(cdf_Chi2_nc)
end

function comp_cdf_Chi2_nc(x::Float64, df::Float64, nc::Float64)::Float64
#// Calculate 1-cdf of noncentral chi-square
#//   parametrized per Johnson & Kotz, SAS, etc. so that cdf_Chi2_nc(x,df,nc) = cdf_gamma_nc(x/2,df/2,nc/2)
#//   if Xi ~ N(Di,1) independent, then sum(Xi,i=1..n) ~ Chi2_nc(n,nc) with nc=sum(Di,i=1..n)
#//   Note that Knusel, Graybill, etc. use a different noncentrality parameter lambda=nc/2
  df = AlterForIntegralChecks_df(df)
  comp_cdf_Chi2_nc = comp_cdf_gamma_nc(x / 2.0, df / 2.0, nc / 2.0)
  comp_cdf_Chi2_nc = GetRidOfMinusZeroes(comp_cdf_Chi2_nc)
end

function inv_Chi2_nc(prob::Float64, df::Float64, nc::Float64)::Float64
   df = AlterForIntegralChecks_df(df)
   inv_Chi2_nc = 2.0 * inv_gamma_nc(prob, df / 2.0, nc / 2.0)
   inv_Chi2_nc = GetRidOfMinusZeroes(inv_Chi2_nc)
end

function comp_inv_Chi2_nc(prob::Float64, df::Float64, nc::Float64)::Float64
   df = AlterForIntegralChecks_df(df)
   comp_inv_Chi2_nc = 2.0 * comp_inv_gamma_nc(prob, df / 2.0, nc / 2.0)
   comp_inv_Chi2_nc = GetRidOfMinusZeroes(comp_inv_Chi2_nc)
end

function ncp_Chi2_nc(prob::Float64, x::Float64, df::Float64)::Float64
   df = AlterForIntegralChecks_df(df)
   ncp_Chi2_nc = 2.0 * ncp_gamma_nc(prob, x / 2.0, df / 2.0)
   ncp_Chi2_nc = GetRidOfMinusZeroes(ncp_Chi2_nc)
end

function comp_ncp_Chi2_nc(prob::Float64, x::Float64, df::Float64)::Float64
   df = AlterForIntegralChecks_df(df)
   comp_ncp_Chi2_nc = 2.0 * comp_ncp_gamma_nc(prob, x / 2.0, df / 2.0)
   comp_ncp_Chi2_nc = GetRidOfMinusZeroes(comp_ncp_Chi2_nc)
end

function  beta_nc1(x::Float64, y::Float64, a::Float64, b::Float64, nc::Float64, ByRef nc_derivative::Float64)::Float64
#y is 1-x but held accurately to avoid possible cancellation errors
   Dim aa::Float64, bb::Float64, nc_dtemp::Float64
   Dim n::Float64, p::Float64, w::Float64, s::Float64, ps::Float64
   Dim result::Float64, term::Float64, ptx::Float64, ptnc::Float64
   bb = (x * nc - 1.0) - a
   if bb < -1E+150
      n = a / bb
      aa = n - nc * x * (n + b / bb)
      n = bb * (1.0 + abs2(1 - (4.0 * aa / bb)))
      n = Int(2.0 * aa * (bb / n))
   else
      aa = a - nc * x * (a + b)
      if (bb < 0.0)
         n = bb - abs2(bb ^ 2 - 4.0 * aa)
         n = Int(2.0 * aa / n)
      else
         n = Int((bb + abs2(bb ^ 2 - 4.0 * aa)) / 2.0)
      end
   end
   if n < 0.0
      n = 0.0
   end
   aa = n + a
   bb = n
   ptnc = poissonTerm(n, nc, nc - n, 0.0)
   ptx = b * binomialTerm(aa, b, x, y, b * x - aa * y, 0.0)  #  (aa + b)*(I(x, aa, b) - I(x, aa + 1, b))
   aa = aa + 1.0
   bb = bb + 1.0
   p = nc / bb
   ps = p
   nc_derivative = ps
   s = x / aa  # (I(x, aa, b) - I(x, aa + 1, b)) / ptx
   w = p
   term = s * w
   result = term
   if ptx > 0
     while (((term > 0.000000000000001 * result) && (p > 1E-16 * w)) || (ps > 1E-16 * nc_derivative))
       s = (aa + b) * s
       aa = aa + 1.0
       bb = bb + 1.0
       p = nc / bb * p
       ps = p * s
       nc_derivative = nc_derivative + ps
       s = x / aa * s # (I(x, aa, b) - I(x, aa + 1, b)) / ptx
       w = w + p
       term = s * w
       result = result + term
     end
     w = w * ptnc
   else
     w = comppoisson(n, nc, nc - n)
   end
   if x > y
      s = compbeta(y, b, a + (bb + 1.0))
   else
      s = beta(x, a + (bb + 1.0), b)
   end
   beta_nc1 = result * ptx * ptnc + s * w
   ps = 1.0
   nc_dtemp = 0.0
   aa = n + a
   bb = n
   p = 1.0
   s = ptx / (aa + b) # I(x, aa, b) - I(x, aa + 1, b)
   if x > y
      w = compbeta(y, b, aa) # I(x, aa, b)
   else
      w = beta(x, aa, b) # I(x, aa, b)
   end
   term = p * w
   result = term
   while bb > 0.0 && (((term > 0.000000000000001 * result) && (s > 1E-16 * w)) || (ps > 1E-16 * nc_dtemp))
       s = aa / x * s
       ps = p * s
       nc_dtemp = nc_dtemp + ps
       p = bb / nc * p
       aa = aa - 1.0
       bb = bb - 1.0
       if bb = 0.0 aa = a
       s = s / (aa + b) # I(x, aa, b) - I(x, aa + 1, b)
       w = w + s # I(x, aa, b)
       term = p * w
       result = result + term
   end
   if n > 0.0
      nc_dtemp = nc_derivative * ptx + nc_dtemp + p * aa / x * s
   elseif b = 0.0
      nc_dtemp = 0.0
   else
      nc_dtemp = binomialTerm(aa, b, x, y, b * x - aa * y, log(b) + log((nc_derivative + aa / (x * (aa + b)))))
   end
   nc_dtemp = nc_dtemp / y
   beta_nc1 = beta_nc1 + result * ptnc + cpoisson(bb - 1.0, nc, nc - bb + 1.0) * w
   if nc_dtemp = 0.0
      nc_derivative = 0.0
   else
      nc_derivative = poissonTerm(n, nc, nc - n, log(nc_dtemp))
   end
end

function  comp_beta_nc1(x::Float64, y::Float64, a::Float64, b::Float64, nc::Float64, ByRef nc_derivative::Float64)::Float64
#y is 1-x but held accurately to avoid possible cancellation errors
   Dim aa::Float64, bb::Float64, nc_dtemp::Float64
   Dim n::Float64, p::Float64, w::Float64, s::Float64, ps::Float64
   Dim result::Float64, term::Float64, ptx::Float64, ptnc::Float64
   bb = (x * nc - 1.0) - a
   if bb < -1E+150
      n = a / bb
      aa = n - nc * x * (n + b / bb)
      n = bb * (1.0 + abs2(1 - (4.0 * aa / bb)))
      n = Int(2.0 * aa * (bb / n))
   else
      aa = a - nc * x * (a + b)
      if (bb < 0.0)
         n = bb - abs2(bb ^ 2 - 4.0 * aa)
         n = Int(2.0 * aa / n)
      else
         n = Int((bb + abs2(bb ^ 2 - 4.0 * aa)) / 2.0)
      end
   end
   if n < 0.0
      n = 0.0
   end
   aa = n + a
   bb = n
   ptnc = poissonTerm(n, nc, nc - n, 0.0)
   ptx = b / (aa + b) * binomialTerm(aa, b, x, y, b * x - aa * y, 0.0) #(1 - I(x, aa + 1, b)) - (1 - I(x, aa, b))
   ps = 1.0
   nc_dtemp = 0.0
   p = 1.0
   s = 1.0
   w = p
   term = 1.0
   result = 0.0
   if ptx > 0
     while bb > 0.0 && (((term > 0.000000000000001 * result) && (p > 1E-16 * w)) || (ps > 1E-16 * nc_dtemp))
       s = aa / x * s
       ps = p * s
       nc_dtemp = nc_dtemp + ps
       p = bb / nc * p
       aa = aa - 1.0
       bb = bb - 1.0
       if bb = 0.0 aa = a
       s = s / (aa + b) # (1 - I(x, aa + 1, b)) - (1 - I(x, aa + 1, b))
       term = s * w
       result = result + term
       w = w + p
     end
     w = w * ptnc
   else
     w = cpoisson(n, nc, nc - n)
   end
   if n > 0.0
      nc_dtemp = (nc_dtemp + p * aa / x * s) * ptx
   elseif a = 0.0 || b = 0.0
      nc_dtemp = 0.0
   else
      nc_dtemp = binomialTerm(aa, b, x, y, b * x - aa * y, log(b) + log(aa / (x * (aa + b))))
   end
   if x > y
      s = beta(y, b, aa)
   else
      s = compbeta(x, aa, b)
   end
   comp_beta_nc1 = result * ptx * ptnc + s * w
   aa = n + a
   bb = n
   p = 1.0
   nc_derivative = 0.0
   s = ptx
   if x > y
      w = beta(y, b, aa) #  1 - I(x, aa, b)
   else
      w = compbeta(x, aa, b) # 1 - I(x, aa, b)
   end
   term = 0.0
   result = term
   Do
       w = w + s # 1 - I(x, aa, b)
       s = (aa + b) * s
       aa = aa + 1.0
       bb = bb + 1.0
       p = nc / bb * p
       ps = p * s
       nc_derivative = nc_derivative + ps
       s = x / aa * s # (1 - I(x, aa + 1, b)) - (1 - I(x, aa, b))
       term = p * w
       result = result + term
   end while (((term > 0.000000000000001 * result) && (s > 1E-16 * w)) || (ps > 1E-16 * nc_derivative))
   nc_dtemp = (nc_derivative + nc_dtemp) / y
   comp_beta_nc1 = comp_beta_nc1 + result * ptnc + comppoisson(bb, nc, nc - bb) * w
   if nc_dtemp = 0.0
      nc_derivative = 0.0
   else
      nc_derivative = poissonTerm(n, nc, nc - n, log(nc_dtemp))
   end
end

function  inv_beta_nc1(prob::Float64, a::Float64, b::Float64, nc::Float64, ByRef oneMinusP::Float64)::Float64
#Uses approx in A&S 26.6.26 for to get initial estimate the modified NR to improve it.
Dim x::Float64, y::Float64, pr::Float64, dif::Float64, temp::Float64
Dim hip::Float64, lop::Float64
Dim hix::Float64, lox::Float64, nc_derivative::Float64
   if (prob > 0.5)
      inv_beta_nc1 = comp_inv_beta_nc1(1.0 - prob, a, b, nc, oneMinusP)
      Exit Function
   end

   lop = 0.0
   hip = 1.0
   lox = 0.0
   hix = 1.0
   pr = exp(-nc)
   if pr > prob
      if 2.0 * prob > pr
         x = invcompbeta(a + cSmall, b, (pr - prob) / pr, oneMinusP)
      else
         x = invbeta(a + cSmall, b, prob / pr, oneMinusP)
      end
      if x = 0.0
         inv_beta_nc1 = 0.0
         Exit Function
      else
         temp = oneMinusP
         y = invbeta(a + nc ^ 2 / (a + 2 * nc), b, prob, oneMinusP)
         oneMinusP = (a + nc) * oneMinusP / (a + nc * (1.0 + y))
         if temp > oneMinusP
            oneMinusP = temp
         else
            x = (a + 2.0 * nc) * y / (a + nc * (1.0 + y))
         end
      end
   else
      y = invbeta(a + nc ^ 2 / (a + 2 * nc), b, prob, oneMinusP)
      x = (a + 2.0 * nc) * y / (a + nc * (1.0 + y))
      oneMinusP = (a + nc) * oneMinusP / (a + nc * (1.0 + y))
      if oneMinusP < cSmall
         oneMinusP = cSmall
         pr = beta_nc1(x, oneMinusP, a, b, nc, nc_derivative)
         if pr < prob
            inv_beta_nc1 = 1.0
            oneMinusP = 0.0
            Exit Function
         end
      end
   end
   Do
      pr = beta_nc1(x, oneMinusP, a, b, nc, nc_derivative)
      if pr < 3E-308 && nc_derivative = 0.0
         hip = oneMinusP
         lox = x
         dif = dif / 2.0
         x = x - dif
         oneMinusP = oneMinusP + dif
      elseif nc_derivative = 0.0
         lop = oneMinusP
         hix = x
         dif = dif / 2.0
         x = x - dif
         oneMinusP = oneMinusP + dif
      else
         if pr < prob
            hip = oneMinusP
            lox = x
         else
            lop = oneMinusP
            hix = x
         end
         dif = -(pr / nc_derivative) * logdif(pr, prob)
         if x > oneMinusP
            if oneMinusP - dif < lop
               dif = (oneMinusP - lop) * 0.9
            elseif oneMinusP - dif > hip
               dif = (oneMinusP - hip) * 0.9
            end
         elseif x + dif < lox
            dif = (lox - x) * 0.9
         elseif x + dif > hix
            dif = (hix - x) * 0.9
         end
         x = x + dif
         oneMinusP = oneMinusP - dif
      end
   end while ((abs(pr - prob) > prob * 0.00000000000001) && (abs(dif) > abs(min(x, oneMinusP)) * 0.0000000001))
   inv_beta_nc1 = x
end

function  comp_inv_beta_nc1(prob::Float64, a::Float64, b::Float64, nc::Float64, ByRef oneMinusP::Float64)::Float64
#Uses approx in A&S 26.6.26 for to get initial estimate the modified NR to improve it.
Dim x::Float64, y::Float64, pr::Float64, dif::Float64, temp::Float64
Dim hip::Float64, lop::Float64
Dim hix::Float64, lox::Float64, nc_derivative::Float64
   if (prob > 0.5)
      comp_inv_beta_nc1 = inv_beta_nc1(1.0 - prob, a, b, nc, oneMinusP)
      Exit Function
   end

   lop = 0.0
   hip = 1.0
   lox = 0.0
   hix = 1.0
   pr = exp(-nc)
   if pr > prob
      if 2.0 * prob > pr
         x = invbeta(a + cSmall, b, (pr - prob) / pr, oneMinusP)
      else
         x = invcompbeta(a + cSmall, b, prob / pr, oneMinusP)
      end
      if oneMinusP < cSmall
         oneMinusP = cSmall
         pr = comp_beta_nc1(x, oneMinusP, a, b, nc, nc_derivative)
         if pr > prob
            comp_inv_beta_nc1 = 1.0
            oneMinusP = 0.0
            Exit Function
         end
      else
         temp = oneMinusP
         y = invcompbeta(a + nc ^ 2 / (a + 2 * nc), b, prob, oneMinusP)
         oneMinusP = (a + nc) * oneMinusP / (a + nc * (1.0 + y))
         if temp < oneMinusP
            oneMinusP = temp
         else
            x = (a + 2.0 * nc) * y / (a + nc * (1.0 + y))
         end
         if oneMinusP < cSmall
            oneMinusP = cSmall
            pr = comp_beta_nc1(x, oneMinusP, a, b, nc, nc_derivative)
            if pr > prob
               comp_inv_beta_nc1 = 1.0
               oneMinusP = 0.0
               Exit Function
            end
         elseif x < cSmall
            x = cSmall
            pr = comp_beta_nc1(x, oneMinusP, a, b, nc, nc_derivative)
            if pr < prob
               comp_inv_beta_nc1 = 0.0
               oneMinusP = 1.0
               Exit Function
            end
         end
      end
   else
      y = invcompbeta(a + nc ^ 2 / (a + 2 * nc), b, prob, oneMinusP)
      x = (a + 2.0 * nc) * y / (a + nc * (1.0 + y))
      oneMinusP = (a + nc) * oneMinusP / (a + nc * (1.0 + y))
      if oneMinusP < cSmall
         oneMinusP = cSmall
         pr = comp_beta_nc1(x, oneMinusP, a, b, nc, nc_derivative)
         if pr > prob
            comp_inv_beta_nc1 = 1.0
            oneMinusP = 0.0
            Exit Function
         end
      elseif x < cSmall
         x = cSmall
         pr = comp_beta_nc1(x, oneMinusP, a, b, nc, nc_derivative)
         if pr < prob
            comp_inv_beta_nc1 = 0.0
            oneMinusP = 1.0
            Exit Function
         end
      end
   end
   dif = x
   Do
      pr = comp_beta_nc1(x, oneMinusP, a, b, nc, nc_derivative)
      if pr < 3E-308 && nc_derivative = 0.0
         lop = oneMinusP
         hix = x
         dif = dif / 2.0
         x = x - dif
         oneMinusP = oneMinusP + dif
      elseif nc_derivative = 0.0
         hip = oneMinusP
         lox = x
         dif = dif / 2.0
         x = x - dif
         oneMinusP = oneMinusP + dif
      else
         if pr < prob
            lop = oneMinusP
            hix = x
         else
            hip = oneMinusP
            lox = x
         end
         dif = (pr / nc_derivative) * logdif(pr, prob)
         if x > oneMinusP
            if oneMinusP - dif < lop
               dif = (oneMinusP - lop) * 0.9
            elseif oneMinusP - dif > hip
               dif = (oneMinusP - hip) * 0.9
            end
         elseif x + dif < lox
            dif = (lox - x) * 0.9
         elseif x + dif > hix
            dif = (hix - x) * 0.9
         end
         x = x + dif
         oneMinusP = oneMinusP - dif
      end
   end while ((abs(pr - prob) > prob * 0.00000000000001) && (abs(dif) > abs(min(x, oneMinusP)) * 0.0000000001))
   comp_inv_beta_nc1 = x
end

function  invBetaLessThanX(prob::Float64, x::Float64, y::Float64, a::Float64, b::Float64)::Float64
   Dim oneMinusP::Float64
   if x >= y
      if invcompbeta(b, a, prob, oneMinusP) >= y * (1.0 - 0.000000000000001)
         invBetaLessThanX = 0.0
      else
         invBetaLessThanX = [#VALUE!]
      end
   elseif invbeta(a, b, prob, oneMinusP) <= x * (1.0 + 0.000000000000001)
      invBetaLessThanX = 0.0
   else
      invBetaLessThanX = [#VALUE!]
   end
end

function  compInvBetaLessThanX(prob::Float64, x::Float64, y::Float64, a::Float64, b::Float64)::Float64
   Dim oneMinusP::Float64
   if x >= y
      if invbeta(b, a, prob, oneMinusP) >= y * (1.0 - 0.000000000000001)
         compInvBetaLessThanX = 0.0
      else
         compInvBetaLessThanX = [#VALUE!]
      end
   elseif invcompbeta(a, b, prob, oneMinusP) <= x * (1.0 + 0.000000000000001)
      compInvBetaLessThanX = 0.0
   else
      compInvBetaLessThanX = [#VALUE!]
   end
end

function  ncp_beta_nc1(prob::Float64, x::Float64, y::Float64, a::Float64, b::Float64)::Float64
#Uses Normal approx for difference of 2 a Negative Binomial and a poisson distributed variable to get initial estimate the modified NR to improve it.
Dim ncp::Float64, pr::Float64, dif::Float64, temp::Float64, deriv::Float64, c::Float64, d::Float64, e::Float64, sqarg::Float64, checked_nc_limit::Bool, checked_0_limit::Bool
Dim hi::Float64, lo::Float64, nc_derivative::Float64
   if (prob > 0.5)
      ncp_beta_nc1 = comp_ncp_beta_nc1(1.0 - prob, x, y, a, b)
      Exit Function
   end

   lo = 0.0
   hi = nc_limit
   checked_0_limit = false
   checked_nc_limit = false
   temp = inv_normal(prob) ^ 2
   c = b * x / y
   d = temp - 2.0 * (a - c)
   if d < 2 * nc_limit
      e = (c - a) ^ 2 - temp * c / y
      sqarg = d ^ 2 - 4 * e
      if sqarg < 0
         ncp = d / 2
      else
         ncp = (d + abs2(sqarg)) / 2
      end
   else
      ncp = nc_limit
   end
   ncp = min(max(0.0, ncp), nc_limit)
   if x > y
      pr = compbeta(y * (1 + ncp / (ncp + a)) / (1 + ncp / (ncp + a) * y), b, a + ncp ^ 2 / (2 * ncp + a))
   else
      pr = beta(x / (1 + ncp / (ncp + a) * y), a + ncp ^ 2 / (2 * ncp + a), b)
   end
   #Debug.Print "ncp_beta_nc1 ncp1 ", ncp, pr
   if ncp = 0.0
      if pr < prob
         ncp_beta_nc1 = invBetaLessThanX(prob, x, y, a, b)
         Exit Function
      else
         checked_0_limit = true
      end
   end
   temp = min(max(0.0, invcompgamma(b * x, prob) / y - a), nc_limit)
   if temp = ncp
      c = pr
   elseif x > y
      c = compbeta(y * (1 + temp / (temp + a)) / (1 + temp / (temp + a) * y), b, a + temp ^ 2 / (2 * temp + a))
   else
      c = beta(x / (1 + temp / (temp + a) * y), a + temp ^ 2 / (2 * temp + a), b)
   end
   #Debug.Print "ncp_beta_nc1 ncp2 ", temp, c
   if temp = 0.0
      if c < prob
         ncp_beta_nc1 = invBetaLessThanX(prob, x, y, a, b)
         Exit Function
      else
         checked_0_limit = true
      end
   end
   if pr * c = 0.0
      ncp = min(ncp, temp)
      pr = max(pr, c)
      if pr = 0.0
         c = compbeta(y, b, a)
         if c < prob
            ncp_beta_nc1 = invBetaLessThanX(prob, x, y, a, b)
            Exit Function
         else
            checked_0_limit = true
         end
      end
   elseif abs(log(pr / prob)) > abs(log(c / prob))
      ncp = temp
      pr = c
   end
   if ncp = 0.0
      if b > 1.000001
         deriv = comp_beta_nc1(x, y, a + 1.0, b - 1.0, ncp, nc_derivative)
         deriv = nc_derivative * y ^ 2 / (b - 1.0)
      else
         deriv = pr - beta_nc1(x, y, a + 1.0, b, ncp, nc_derivative)
      end
      if deriv = 0.0
         ncp = nc_limit
      else
         ncp = (pr - prob) / deriv
         if ncp >= nc_limit
            ncp = (pr / deriv) * logdif(pr, prob)
         end
      end
   else
      if ncp = nc_limit
         if pr > prob
            ncp_beta_nc1 = [#VALUE!]
            Exit Function
         else
            checked_nc_limit = true
         end
      end
      if pr > 0
         temp = ncp * 0.999999 #Use numerical derivative on approximation since cheap compared to evaluating non-central beta
         if x > y
            c = compbeta(y * (1.0 + temp / (temp + a)) / (1 + temp / (temp + a) * y), b, a + temp ^ 2 / (2 * temp + a))
         else
            c = beta(x / (1 + temp / (temp + a) * y), a + temp ^ 2 / (2 * temp + a), b)
         end
         if pr <> c
            dif = (0.000001 * ncp * pr / (pr - c)) * logdif(pr, prob)
            if ncp - dif < 0.0
               ncp = ncp / 2.0
            elseif ncp - dif > nc_limit
               ncp = (ncp + nc_limit) / 2.0
            else
               ncp = ncp - dif
            end
         end
      else
         ncp = ncp / 2.0
      end
   end
   dif = ncp
   Do
      pr = beta_nc1(x, y, a, b, ncp, nc_derivative)
      #Debug.Print ncp, pr, prob
      if b > 1.000001
         deriv = beta_nc1(x, y, a + 1.0, b - 1.0, ncp, nc_derivative)
         deriv = nc_derivative * y ^ 2 / (b - 1.0)
      else
         deriv = pr - beta_nc1(x, y, a + 1.0, b, ncp, nc_derivative)
      end
      if pr < 3E-308 && deriv = 0.0
         hi = ncp
         dif = dif / 2.0
         ncp = ncp - dif
      elseif deriv = 0.0
         lo = ncp
         dif = dif / 2.0
         ncp = ncp - dif
      else
         if pr < prob
            hi = ncp
         else
            lo = ncp
         end
         dif = (pr / deriv) * logdif(pr, prob)
         if ncp + dif < lo
            dif = (lo - ncp) / 2.0
            if Not checked_0_limit && (lo = 0.0)
               temp = cdf_beta_nc(x, a, b, lo)
               if temp < prob
                  ncp_beta_nc1 = invBetaLessThanX(prob, x, y, a, b)
                  Exit Function
               else
                  checked_0_limit = true
               end
            end
         elseif ncp + dif > hi
            dif = (hi - ncp) / 2.0
            if Not checked_nc_limit && (hi = nc_limit)
               temp = cdf_beta_nc(x, a, b, hi)
               if temp > prob
                  ncp_beta_nc1 = [#VALUE!]
                  Exit Function
               else
                  checked_nc_limit = true
               end
            end
         end
         ncp = ncp + dif
      end
   end while ((abs(pr - prob) > prob * 0.00000000000001) && (abs(dif) > abs(ncp) * 0.0000000001))
   ncp_beta_nc1 = ncp
   #Debug.Print "ncp_beta_nc1", ncp_beta_nc1
end

function  comp_ncp_beta_nc1(prob::Float64, x::Float64, y::Float64, a::Float64, b::Float64)::Float64
#Uses Normal approx for difference of 2 a Negative Binomial and a poisson distributed variable to get initial estimate the modified NR to improve it.
Dim ncp::Float64, pr::Float64, dif::Float64, temp::Float64, deriv::Float64, c::Float64, d::Float64, e::Float64, sqarg::Float64, checked_nc_limit::Bool, checked_0_limit::Bool
Dim hi::Float64, lo::Float64, nc_derivative::Float64
   if (prob > 0.5)
      comp_ncp_beta_nc1 = ncp_beta_nc1(1.0 - prob, x, y, a, b)
      Exit Function
   end

   lo = 0.0
   hi = nc_limit
   checked_0_limit = false
   checked_nc_limit = false
   temp = inv_normal(prob) ^ 2
   c = b * x / y
   d = temp - 2.0 * (a - c)
   if d < 4 * nc_limit
      sqarg = d ^ 2 - 4 * e
      if sqarg < 0
         ncp = d / 2
      else
         ncp = (d - abs2(sqarg)) / 2
      end
   else
      ncp = 0.0
   end
   ncp = min(max(0.0, ncp), nc_limit)
   if x > y
      pr = beta(y * (1 + ncp / (ncp + a)) / (1 + ncp / (ncp + a) * y), b, a + ncp ^ 2 / (2 * ncp + a))
   else
      pr = compbeta(x / (1 + ncp / (ncp + a) * y), a + ncp ^ 2 / (2 * ncp + a), b)
   end
   #Debug.Print "comp_ncp_beta_nc1 ncp1 ", ncp, pr
   if ncp = 0.0
      if pr > prob
         comp_ncp_beta_nc1 = compInvBetaLessThanX(prob, x, y, a, b)
         Exit Function
      else
         checked_0_limit = true
      end
   end
   temp = min(max(0.0, invgamma(b * x, prob) / y - a), nc_limit)
   if temp = ncp
      c = pr
   elseif x > y
      c = beta(y * (1 + temp / (temp + a)) / (1 + temp / (temp + a) * y), b, a + temp ^ 2 / (2 * temp + a))
   else
      c = compbeta(x / (1 + temp / (temp + a) * y), a + temp ^ 2 / (2 * temp + a), b)
   end
   #Debug.Print "comp_ncp_beta_nc1 ncp2 ", temp, c
   if temp = 0.0
      if c > prob
         comp_ncp_beta_nc1 = compInvBetaLessThanX(prob, x, y, a, b)
         Exit Function
      else
         checked_0_limit = true
      end
   end
   if pr * c = 0.0
      ncp = max(ncp, temp)
      pr = max(pr, c)
   elseif abs(log(pr / prob)) > abs(log(c / prob))
      ncp = temp
      pr = c
   end
   if ncp = 0.0
      if pr > prob
         comp_ncp_beta_nc1 = compInvBetaLessThanX(prob, x, y, a, b)
         Exit Function
      else
         if b > 1.000001
            deriv = beta_nc1(x, y, a + 1.0, b - 1.0, 0.0, nc_derivative)
            deriv = nc_derivative * y ^ 2 / (b - 1.0)
         else
            deriv = comp_beta_nc1(x, y, a + 1.0, b, 0.0, nc_derivative) - pr
         end
         if deriv = 0.0
            ncp = nc_limit
         else
            ncp = (prob - pr) / deriv
            if ncp >= nc_limit
               ncp = -(pr / deriv) * logdif(pr, prob)
            end
         end
         checked_0_limit = true
      end
   else
      if ncp = nc_limit
         if pr < prob
            comp_ncp_beta_nc1 = [#VALUE!]
            Exit Function
         else
            checked_nc_limit = true
         end
      end
      if pr > 0
         temp = ncp * 0.999999 #Use numerical derivative on approximation since cheap compared to evaluating non-central beta
         if x > y
            c = beta(y * (1.0 + temp / (temp + a)) / (1 + temp / (temp + a) * y), b, a + temp ^ 2 / (2 * temp + a))
         else
            c = compbeta(x / (1 + temp / (temp + a) * y), a + temp ^ 2 / (2 * temp + a), b)
         end
         if pr <> c
            dif = -(0.000001 * ncp * pr / (pr - c)) * logdif(pr, prob)
            if ncp + dif < 0
               ncp = ncp / 2
            elseif ncp + dif > nc_limit
               ncp = (ncp + nc_limit) / 2
            else
               ncp = ncp + dif
            end
         end
      else
         ncp = (nc_limit + ncp) / 2.0
      end
   end
   dif = ncp
   Do
      pr = comp_beta_nc1(x, y, a, b, ncp, nc_derivative)
      #Debug.Print ncp, pr, prob
      if b > 1.000001
         deriv = beta_nc1(x, y, a + 1.0, b - 1.0, ncp, nc_derivative)
         deriv = nc_derivative * y ^ 2 / (b - 1.0)
      else
         deriv = comp_beta_nc1(x, y, a + 1.0, b, ncp, nc_derivative) - pr
      end
      if pr < 3E-308 && deriv = 0.0
         lo = ncp
         dif = dif / 2.0
         ncp = ncp + dif
      elseif deriv = 0.0
         hi = ncp
         dif = dif / 2.0
         ncp = ncp - dif
      else
         if pr < prob
            lo = ncp
         else
            hi = ncp
         end
         dif = -(pr / deriv) * logdif(pr, prob)
         if ncp + dif < lo
            dif = (lo - ncp) / 2.0
            if Not checked_0_limit && (lo = 0.0)
               temp = comp_cdf_beta_nc(x, a, b, lo)
               if temp > prob
                  comp_ncp_beta_nc1 = compInvBetaLessThanX(prob, x, y, a, b)
                  Exit Function
               else
                  checked_0_limit = true
               end
            end
         elseif ncp + dif > hi
            dif = (hi - ncp) / 2.0
            if Not checked_nc_limit && (hi = nc_limit)
               temp = comp_cdf_beta_nc(x, a, b, hi)
               if temp < prob
                  comp_ncp_beta_nc1 = [#VALUE!]
                  Exit Function
               else
                  checked_nc_limit = true
               end
            end
         end
         ncp = ncp + dif
      end
   end while ((abs(pr - prob) > prob * 0.00000000000001) && (abs(dif) > abs(ncp) * 0.0000000001))
   comp_ncp_beta_nc1 = ncp
   #Debug.Print "comp_ncp_beta_nc1", comp_ncp_beta_nc1
end

function pdf_beta_nc(x::Float64, shape_param1::Float64, shape_param2::Float64, nc_param::Float64)::Float64
  if (shape_param1 < 0.0) || (shape_param2 < 0.0) || (nc_param < 0.0) || (nc_param > nc_limit) || ((shape_param1 = 0.0) && (shape_param2 = 0.0))
     pdf_beta_nc = [#VALUE!]
  elseif (x < 0.0 || x > 1.0)
     pdf_beta_nc = 0.0
  elseif (x = 0.0 || nc_param = 0.0)
     pdf_beta_nc = exp(-nc_param) * pdf_beta(x, shape_param1, shape_param2)
  elseif (x = 1.0 && shape_param2 = 1.0)
     pdf_beta_nc = shape_param1 + nc_param
  elseif (x = 1.0)
     pdf_beta_nc = pdf_beta(x, shape_param1, shape_param2)
  else
     Dim nc_derivative::Float64
     if (shape_param1 < 1.0 || x * shape_param2 <= (1.0 - x) * (shape_param1 + nc_param))
        pdf_beta_nc = beta_nc1(x, 1.0 - x, shape_param1, shape_param2, nc_param, nc_derivative)
     else
        pdf_beta_nc = comp_beta_nc1(x, 1.0 - x, shape_param1, shape_param2, nc_param, nc_derivative)
     end
     pdf_beta_nc = nc_derivative
  end
  pdf_beta_nc = GetRidOfMinusZeroes(pdf_beta_nc)
end

function cdf_beta_nc(x::Float64, shape_param1::Float64, shape_param2::Float64, nc_param::Float64)::Float64
  Dim nc_derivative::Float64
  if (shape_param1 < 0.0) || (shape_param2 < 0.0) || (nc_param < 0.0) || (nc_param > nc_limit) || ((shape_param1 = 0.0) && (shape_param2 = 0.0))
     cdf_beta_nc = [#VALUE!]
  elseif (x < 0.0)
     cdf_beta_nc = 0.0
  elseif (x >= 1.0)
     cdf_beta_nc = 1.0
  elseif (x = 0.0 && shape_param1 = 0.0)
     cdf_beta_nc = exp(-nc_param)
  elseif (x = 0.0)
     cdf_beta_nc = 0.0
  elseif (nc_param = 0.0)
     cdf_beta_nc = beta(x, shape_param1, shape_param2)
  elseif (shape_param1 < 1.0 || x * shape_param2 <= (1.0 - x) * (shape_param1 + nc_param))
     cdf_beta_nc = beta_nc1(x, 1.0 - x, shape_param1, shape_param2, nc_param, nc_derivative)
  else
     cdf_beta_nc = 1.0 - comp_beta_nc1(x, 1.0 - x, shape_param1, shape_param2, nc_param, nc_derivative)
  end
  cdf_beta_nc = GetRidOfMinusZeroes(cdf_beta_nc)
end

function comp_cdf_beta_nc(x::Float64, shape_param1::Float64, shape_param2::Float64, nc_param::Float64)::Float64
  Dim nc_derivative::Float64
  if (shape_param1 < 0.0) || (shape_param2 < 0.0) || (nc_param < 0.0) || (nc_param > nc_limit) || ((shape_param1 = 0.0) && (shape_param2 = 0.0))
     comp_cdf_beta_nc = [#VALUE!]
  elseif (x < 0.0)
     comp_cdf_beta_nc = 1.0
  elseif (x >= 1.0)
     comp_cdf_beta_nc = 0.0
  elseif (x = 0.0 && shape_param1 = 0.0)
     comp_cdf_beta_nc = -expm1(-nc_param)
  elseif (x = 0.0)
     comp_cdf_beta_nc = 1.0
  elseif (nc_param = 0.0)
     comp_cdf_beta_nc = compbeta(x, shape_param1, shape_param2)
  elseif (shape_param1 < 1.0 || x * shape_param2 >= (1.0 - x) * (shape_param1 + nc_param))
     comp_cdf_beta_nc = comp_beta_nc1(x, 1.0 - x, shape_param1, shape_param2, nc_param, nc_derivative)
  else
     comp_cdf_beta_nc = 1.0 - beta_nc1(x, 1.0 - x, shape_param1, shape_param2, nc_param, nc_derivative)
  end
  comp_cdf_beta_nc = GetRidOfMinusZeroes(comp_cdf_beta_nc)
end

function inv_beta_nc(prob::Float64, shape_param1::Float64, shape_param2::Float64, nc_param::Float64)::Float64
  Dim oneMinusP::Float64
  if (shape_param1 < 0.0) || (shape_param2 <= 0.0) || (nc_param < 0.0) || (nc_param > nc_limit) || (prob < 0.0) || (prob > 1.0)
     inv_beta_nc = [#VALUE!]
  elseif (prob = 0.0 || shape_param1 = 0.0 && prob <= exp(-nc_param))
     inv_beta_nc = 0.0
  elseif (prob = 1.0)
     inv_beta_nc = 1.0
  elseif (nc_param = 0.0)
     inv_beta_nc = invbeta(shape_param1, shape_param2, prob, oneMinusP)
  else
     inv_beta_nc = inv_beta_nc1(prob, shape_param1, shape_param2, nc_param, oneMinusP)
  end
  inv_beta_nc = GetRidOfMinusZeroes(inv_beta_nc)
end

function comp_inv_beta_nc(prob::Float64, shape_param1::Float64, shape_param2::Float64, nc_param::Float64)::Float64
  Dim oneMinusP::Float64
  if (shape_param1 < 0.0) || (shape_param2 <= 0.0) || (nc_param < 0.0) || (nc_param > nc_limit) || (prob < 0.0) || (prob > 1.0)
     comp_inv_beta_nc = [#VALUE!]
  elseif (prob = 1.0 || shape_param1 = 0.0 && prob >= -expm1(-nc_param))
     comp_inv_beta_nc = 0.0
  elseif (prob = 0.0)
     comp_inv_beta_nc = 1.0
  elseif (nc_param = 0.0)
     comp_inv_beta_nc = invcompbeta(shape_param1, shape_param2, prob, oneMinusP)
  else
     comp_inv_beta_nc = comp_inv_beta_nc1(prob, shape_param1, shape_param2, nc_param, oneMinusP)
  end
  comp_inv_beta_nc = GetRidOfMinusZeroes(comp_inv_beta_nc)
end

function ncp_beta_nc(prob::Float64, x::Float64, shape_param1::Float64, shape_param2::Float64)::Float64
  if (shape_param1 < 0.0) || (shape_param2 <= 0.0) || (x < 0.0) || (x >= 1.0) || (prob <= 0.0) || (prob > 1.0)
     ncp_beta_nc = [#VALUE!]
  elseif (x = 0.0 && shape_param1 = 0.0)
     ncp_beta_nc = -log(prob)
  elseif (x = 0.0 || prob = 1.0)
     ncp_beta_nc = [#VALUE!]
  else
     ncp_beta_nc = ncp_beta_nc1(prob, x, 1.0 - x, shape_param1, shape_param2)
  end
  ncp_beta_nc = GetRidOfMinusZeroes(ncp_beta_nc)
end

function comp_ncp_beta_nc(prob::Float64, x::Float64, shape_param1::Float64, shape_param2::Float64)::Float64
  if (shape_param1 < 0.0) || (shape_param2 <= 0.0) || (x < 0.0) || (x >= 1.0) || (prob < 0.0) || (prob >= 1.0)
     comp_ncp_beta_nc = [#VALUE!]
  elseif (x = 0.0 && shape_param1 = 0.0)
     comp_ncp_beta_nc = -log0(-prob)
  elseif (x = 0.0 || prob = 0.0)
     comp_ncp_beta_nc = [#VALUE!]
  else
     comp_ncp_beta_nc = comp_ncp_beta_nc1(prob, x, 1.0 - x, shape_param1, shape_param2)
  end
  comp_ncp_beta_nc = GetRidOfMinusZeroes(comp_ncp_beta_nc)
end

function pdf_fdist_nc(x::Float64, df1::Float64, df2::Float64, nc::Float64)::Float64
   df1 = AlterForIntegralChecks_df(df1)
   df2 = AlterForIntegralChecks_df(df2)
   if (df1 <= 0.0 || df2 <= 0.0 || (nc < 0.0) || (nc > 2.0 * nc_limit))
      pdf_fdist_nc = [#VALUE!]
   elseif (x < 0.0)
      pdf_fdist_nc = 0.0
   elseif (x = 0.0 || nc = 0.0)
      pdf_fdist_nc = exp(-nc / 2.0) * pdf_fdist(x, df1, df2)
   else
      Dim p::Float64, q::Float64, nc_derivative::Float64
      if x > 1.0
         q = df2 / x
         p = q + df1
         q = q / p
         p = df1 / p
      else
         p = df1 * x
         q = df2 + p
         p = p / q
         q = df2 / q
      end
      if (df1 < 1.0 || p * df2 <= q * (df1 + nc))
         pdf_fdist_nc = beta_nc1(p, q, df1 / 2.0, df2 / 2.0, nc / 2.0, nc_derivative)
      else
         pdf_fdist_nc = comp_beta_nc1(p, q, df1 / 2.0, df2 / 2.0, nc / 2.0, nc_derivative)
      end
      pdf_fdist_nc = (nc_derivative * q) * (df1 * q / df2)
   end
   pdf_fdist_nc = GetRidOfMinusZeroes(pdf_fdist_nc)
end

function cdf_fdist_nc(x::Float64, df1::Float64, df2::Float64, nc::Float64)::Float64
   df1 = AlterForIntegralChecks_df(df1)
   df2 = AlterForIntegralChecks_df(df2)
   if (df1 <= 0.0 || df2 <= 0.0 || (nc < 0.0) || (nc > 2.0 * nc_limit))
      cdf_fdist_nc = [#VALUE!]
   elseif (x <= 0.0)
      cdf_fdist_nc = 0.0
   else
      Dim p::Float64, q::Float64, nc_derivative::Float64
      if x > 1.0
         q = df2 / x
         p = q + df1
         q = q / p
         p = df1 / p
      else
         p = df1 * x
         q = df2 + p
         p = p / q
         q = df2 / q
      end
      #if p < cSmall && x <> 0.0 || q < cSmall
      #   cdf_fdist_nc = [#VALUE!]
      #   Exit Function
      #end
      df2 = df2 / 2.0
      df1 = df1 / 2.0
      nc = nc / 2.0
      if (nc = 0.0 && p <= q)
         cdf_fdist_nc = beta(p, df1, df2)
      elseif (nc = 0.0)
         cdf_fdist_nc = compbeta(q, df2, df1)
      elseif (df1 < 1.0 || p * df2 <= q * (df1 + nc))
         cdf_fdist_nc = beta_nc1(p, q, df1, df2, nc, nc_derivative)
      else
         cdf_fdist_nc = 1.0 - comp_beta_nc1(p, q, df1, df2, nc, nc_derivative)
      end
   end
   cdf_fdist_nc = GetRidOfMinusZeroes(cdf_fdist_nc)
end

function comp_cdf_fdist_nc(x::Float64, df1::Float64, df2::Float64, nc::Float64)::Float64
   df1 = AlterForIntegralChecks_df(df1)
   df2 = AlterForIntegralChecks_df(df2)
   if (df1 <= 0.0 || df2 <= 0.0 || (nc < 0.0) || (nc > 2.0 * nc_limit))
      comp_cdf_fdist_nc = [#VALUE!]
   elseif (x <= 0.0)
      comp_cdf_fdist_nc = 1.0
   else
      Dim p::Float64, q::Float64, nc_derivative::Float64
      if x > 1.0
         q = df2 / x
         p = q + df1
         q = q / p
         p = df1 / p
      else
         p = df1 * x
         q = df2 + p
         p = p / q
         q = df2 / q
      end
      #if p < cSmall && x <> 0.0 || q < cSmall
      #   comp_cdf_fdist_nc = [#VALUE!]
      #   Exit Function
      #end
      df2 = df2 / 2.0
      df1 = df1 / 2.0
      nc = nc / 2.0
      if (nc = 0.0 && p <= q)
         comp_cdf_fdist_nc = compbeta(p, df1, df2)
      elseif (nc = 0.0)
         comp_cdf_fdist_nc = beta(q, df2, df1)
      elseif (df1 < 1.0 || p * df2 >= q * (df1 + nc))
         comp_cdf_fdist_nc = comp_beta_nc1(p, q, df1, df2, nc, nc_derivative)
      else
         comp_cdf_fdist_nc = 1.0 - beta_nc1(p, q, df1, df2, nc, nc_derivative)
      end
   end
   comp_cdf_fdist_nc = GetRidOfMinusZeroes(comp_cdf_fdist_nc)
end

function inv_fdist_nc(prob::Float64, df1::Float64, df2::Float64, nc::Float64)::Float64
   df1 = AlterForIntegralChecks_df(df1)
   df2 = AlterForIntegralChecks_df(df2)
   if (df1 <= 0.0 || df2 <= 0.0 || (nc < 0.0) || (nc > 2.0 * nc_limit) || prob < 0.0 || prob >= 1.0)
      inv_fdist_nc = [#VALUE!]
   elseif (prob = 0.0)
      inv_fdist_nc = 0.0
   else
      Dim temp::Float64, oneMinusP::Float64
      df1 = df1 / 2.0
      df2 = df2 / 2.0
      if nc = 0.0
         temp = invbeta(df1, df2, prob, oneMinusP)
      else
         temp = inv_beta_nc1(prob, df1, df2, nc / 2.0, oneMinusP)
      end
      inv_fdist_nc = df2 * temp / (df1 * oneMinusP)
      #if oneMinusP < cSmall inv_fdist_nc = [#VALUE!]
   end
   inv_fdist_nc = GetRidOfMinusZeroes(inv_fdist_nc)
end

function comp_inv_fdist_nc(prob::Float64, df1::Float64, df2::Float64, nc::Float64)::Float64
   df1 = AlterForIntegralChecks_df(df1)
   df2 = AlterForIntegralChecks_df(df2)
   if (df1 <= 0.0 || df2 <= 0.0 || (nc < 0.0) || (nc > 2.0 * nc_limit) || prob <= 0.0 || prob > 1.0)
      comp_inv_fdist_nc = [#VALUE!]
   elseif (prob = 1.0)
      comp_inv_fdist_nc = 0.0
   else
      Dim temp::Float64, oneMinusP::Float64
      df1 = df1 / 2.0
      df2 = df2 / 2.0
      if nc = 0.0
         temp = invcompbeta(df1, df2, prob, oneMinusP)
      else
         temp = comp_inv_beta_nc1(prob, df1, df2, nc / 2.0, oneMinusP)
      end
      comp_inv_fdist_nc = df2 * temp / (df1 * oneMinusP)
      #if oneMinusP < cSmall comp_inv_fdist_nc = [#VALUE!]
   end
   comp_inv_fdist_nc = GetRidOfMinusZeroes(comp_inv_fdist_nc)
end

function ncp_fdist_nc(prob::Float64, x::Float64, df1::Float64, df2::Float64)::Float64
  df1 = AlterForIntegralChecks_df(df1)
  df2 = AlterForIntegralChecks_df(df2)
  if (df1 <= 0.0) || (df2 <= 0.0) || (x <= 0.0) || (prob <= 0.0) || (prob >= 1.0)
     ncp_fdist_nc = [#VALUE!]
  else
     Dim p::Float64, q::Float64
     if x > 1.0
        q = df2 / x
        p = q + df1
        q = q / p
        p = df1 / p
     else
        p = df1 * x
        q = df2 + p
        p = p / q
        q = df2 / q
     end
     df2 = df2 / 2.0
     df1 = df1 / 2.0
     ncp_fdist_nc = ncp_beta_nc1(prob, p, q, df1, df2) * 2.0
  end
  ncp_fdist_nc = GetRidOfMinusZeroes(ncp_fdist_nc)
end

function comp_ncp_fdist_nc(prob::Float64, x::Float64, df1::Float64, df2::Float64)::Float64
  df1 = AlterForIntegralChecks_df(df1)
  df2 = AlterForIntegralChecks_df(df2)
  if (df1 <= 0.0) || (df2 <= 0.0) || (x <= 0.0) || (prob <= 0.0) || (prob >= 1.0)
     comp_ncp_fdist_nc = [#VALUE!]
  else
     Dim p::Float64, q::Float64
     if x > 1.0
        q = df2 / x
        p = q + df1
        q = q / p
        p = df1 / p
     else
        p = df1 * x
        q = df2 + p
        p = p / q
        q = df2 / q
     end
     df1 = df1 / 2.0
     df2 = df2 / 2.0
     comp_ncp_fdist_nc = comp_ncp_beta_nc1(prob, p, q, df1, df2) * 2.0
  end
  comp_ncp_fdist_nc = GetRidOfMinusZeroes(comp_ncp_fdist_nc)
end

function  t_nc1(t::Float64, df::Float64, nct::Float64, ByRef nc_derivative::Float64)::Float64
#y is 1-x but held accurately to avoid possible cancellation errors
#nc_derivative holds t * derivative
   Dim aa::Float64, bb::Float64, nc_dtemp::Float64
   Dim n::Float64, p::Float64, q::Float64, w::Float64, V::Float64, r::Float64, s::Float64, ps::Float64
   Dim result1::Float64, result2::Float64, term1::Float64, term2::Float64, ptnc::Float64, qtnc::Float64, ptx::Float64, qtx::Float64
   Dim a::Float64, b::Float64, x::Float64, y::Float64, nc::Float64
   Dim save_result1::Float64, save_result2::Float64, phi::Float64, vScale::Float64
   phi = cnormal(-abs(nct))
   a = 0.5
   b = df / 2.0
   if abs(t) >= min(1.0, df)
      y = df / t
      x = t + y
      y = y / x
      x = t / x
   else
      x = t * t
      y = df + x
      x = x / y
      y = df / y
   end
   if y < cSmall
      t_nc1 = [#VALUE!]
      Exit Function
   end
   nc = nct * nct / 2.0
   aa = a - nc * x * (a + b)
   bb = (x * nc - 1.0) - a
   if (bb < 0.0)
      n = bb - abs2(bb ^ 2 - 4.0 * aa)
      n = Int(2.0 * aa / n)
   else
      n = Int((bb + abs2(bb ^ 2 - 4.0 * aa)) / 2.0)
   end
   if n < 0.0
      n = 0.0
   end
   aa = n + a
   bb = n + 0.5
   qtnc = poissonTerm(bb, nc, nc - bb, 0.0)
   bb = n
   ptnc = poissonTerm(bb, nc, nc - bb, 0.0)
   ptx = binomialTerm(aa, b, x, y, b * x - aa * y, 0.0) / (aa + b) #(I(x, aa, b) - I(x, aa+1, b))/b
   qtx = binomialTerm(aa + 0.5, b, x, y, b * x - (aa + 0.5) * y, 0.0) / (aa + b + 0.5) #(I(x, aa+1/2, b) - I(x, aa+3/2, b))/b
   if b > 1.0
      ptx = b * ptx
      qtx = b * qtx
   end
   vScale = max(ptx, qtx)
   if ptx = vScale
      s = 1.0
   else
      s = ptx / vScale
   end
   if qtx = vScale
      r = 1.0
   else
      r = qtx / vScale
   end
   s = (aa + b) * s
   r = (aa + b + 0.5) * r
   aa = aa + 1.0
   bb = bb + 1.0
   p = nc / bb * ptnc
   q = nc / (bb + 0.5) * qtnc
   ps = p * s + q * r
   nc_derivative = ps
   s = x / aa * s  # I(x, aa, b) - I(x, aa+1, b)
   r = x / (aa + 0.5) * r # I(x, aa+1/2, b) - I(x, aa+3/2, b)
   w = p
   V = q
   term1 = s * w
   term2 = r * V
   result1 = term1
   result2 = term2
   while ((((term1 + term2) > 0.000000000000001 * (result1 + result2)) && (p > 1E-16 * w)) || (ps > 1E-16 * nc_derivative))
       s = (aa + b) * s
       r = (aa + b + 0.5) * r
       aa = aa + 1.0
       bb = bb + 1.0
       p = nc / bb * p
       q = nc / (bb + 0.5) * q
       ps = p * s + q * r
       nc_derivative = nc_derivative + ps
       s = x / aa * s # I(x, aa, b) - I(x, aa+1, b)
       r = x / (aa + 0.5) * r # I(x, aa+1/2, b) - I(x, aa+3/2, b)
       w = w + p
       V = V + q
       term1 = s * w
       term2 = r * V
       result1 = result1 + term1
       result2 = result2 + term2
   end
   if x > y
      s = compbeta(y, b, a + (bb + 1.0))
      r = compbeta(y, b, a + (bb + 1.5))
   else
      s = beta(x, a + (bb + 1.0), b)
      r = beta(x, a + (bb + 1.5), b)
   end
   nc_derivative = x * nc_derivative * vScale
   if b <= 1.0 vScale = vScale * b
   save_result1 = result1 * vScale + s * w
   save_result2 = result2 * vScale + r * V

   ps = 1.0
   nc_dtemp = 0.0
   aa = n + a
   bb = n
   vScale = max(ptnc, qtnc)
   if ptnc = vScale
      p = 1.0
   else
      p = ptnc / vScale
   end
   if qtnc = vScale
      q = 1.0
   else
      q = qtnc / vScale
   end
   s = ptx # I(x, aa, b) - I(x, aa+1, b)
   r = qtx # I(x, aa+1/2, b) - I(x, aa+3/2, b)
   if x > y
      w = compbeta(y, b, aa) # I(x, aa, b)
      V = compbeta(y, b, aa + 0.5) # I(x, aa+1/2, b)
   else
      w = beta(x, aa, b) # I(x, aa, b)
      V = beta(x, aa + 0.5, b) # I(x, aa+1/2, b)
   end
   term1 = p * w
   term2 = q * V
   result1 = term1
   result2 = term2
   while bb > 0.0 && ((((term1 + term2) > 0.000000000000001 * (result1 + result2)) && (s > 1E-16 * w)) || (ps > 1E-16 * nc_dtemp))
       s = aa / x * s
       r = (aa + 0.5) / x * r
       ps = p * s + q * r
       nc_dtemp = nc_dtemp + ps
       p = bb / nc * p
       q = (bb + 0.5) / nc * q
       aa = aa - 1.0
       bb = bb - 1.0
       if bb = 0.0 aa = a
       s = s / (aa + b) # I(x, aa, b) - I(x, aa+1, b)
       r = r / (aa + b + 0.5) # I(x, aa+1/2, b) - I(x, aa+3/2, b)
       if b > 1.0
          w = w + s # I(x, aa, b)
          V = V + r # I(x, aa+0.5, b)
       else
          w = w + b * s
          V = V + b * r
       end
       term1 = p * w
       term2 = q * V
       result1 = result1 + term1
       result2 = result2 + term2
   end
   nc_dtemp = x * nc_dtemp + p * aa * s + q * (aa + 0.5) * r
   p = cpoisson(bb - 1.0, nc, nc - bb + 1.0)
   q = cpoisson(bb - 0.5, nc, nc - bb + 0.5) - 2.0 * phi
   result1 = save_result1 + result1 * vScale + p * w
   result2 = save_result2 + result2 * vScale + q * V
   if t > 0.0
      t_nc1 = phi + 0.5 * (result1 + result2)
      nc_derivative = nc_derivative + nc_dtemp * vScale
   else
      t_nc1 = phi - 0.5 * (result1 - result2)
   end
end

function  comp_t_nc1(t::Float64, df::Float64, nct::Float64, ByRef nc_derivative::Float64)::Float64
#y is 1-x but held accurately to avoid possible cancellation errors
#nc_derivative holds t * derivative
   Dim aa::Float64, bb::Float64, nc_dtemp::Float64
   Dim n::Float64, p::Float64, q::Float64, w::Float64, V::Float64, r::Float64, s::Float64, ps::Float64
   Dim result1::Float64, result2::Float64, term1::Float64, term2::Float64, ptnc::Float64, qtnc::Float64, ptx::Float64, qtx::Float64
   Dim a::Float64, b::Float64, x::Float64, y::Float64, nc::Float64
   Dim save_result1::Float64, save_result2::Float64, vScale::Float64
   a = 0.5
   b = df / 2.0
   if abs(t) >= min(1.0, df)
      y = df / t
      x = t + y
      y = y / x
      x = t / x
   else
      x = t * t
      y = df + x
      x = x / y
      y = df / y
   end
   if y < cSmall
      comp_t_nc1 = [#VALUE!]
      Exit Function
   end
   nc = nct * nct / 2.0
   aa = a - nc * x * (a + b)
   bb = (x * nc - 1.0) - a
   if (bb < 0.0)
      n = bb - abs2(bb ^ 2 - 4.0 * aa)
      n = Int(2.0 * aa / n)
   else
      n = Int((bb + abs2(bb ^ 2 - 4.0 * aa)) / 2)
   end
   if n < 0.0
      n = 0.0
   end
   aa = n + a
   bb = n + 0.5
   qtnc = poissonTerm(bb, nc, nc - bb, 0.0)
   bb = n
   ptnc = poissonTerm(bb, nc, nc - bb, 0.0)
   ptx = binomialTerm(aa, b, x, y, b * x - aa * y, 0.0) / (aa + b) #((1 - I(x, aa+1, b)) - (1 - I(x, aa, b)))/b
   qtx = binomialTerm(aa + 0.5, b, x, y, b * x - (aa + 0.5) * y, 0.0) / (aa + b + 0.5) #((1 - I(x, aa+3/2, b)) - (1 - I(x, aa+1/2, b)))/b
   if b > 1.0
      ptx = b * ptx
      qtx = b * qtx
   end
   vScale = max(ptnc, qtnc)
   if ptnc = vScale
      p = 1.0
   else
      p = ptnc / vScale
   end
   if qtnc = vScale
      q = 1.0
   else
      q = qtnc / vScale
   end
   nc_derivative = 0.0
   s = ptx
   r = qtx
   if x > y
      V = beta(y, b, aa + 0.5) #  1 - I(x, aa+1/2, b)
      w = beta(y, b, aa) #  1 - I(x, aa, b)
   else
      V = compbeta(x, aa + 0.5, b) # 1 - I(x, aa+1/2, b)
      w = compbeta(x, aa, b) # 1 - I(x, aa, b)
   end
   term1 = 0.0
   term2 = 0.0
   result1 = term1
   result2 = term2
   Do
       if b > 1.0
          w = w + s # 1 - I(x, aa, b)
          V = V + r # 1 - I(x, aa+1/2, b)
       else
          w = w + b * s
          V = V + b * r
       end
       s = (aa + b) * s
       r = (aa + b + 0.5) * r
       aa = aa + 1.0
       bb = bb + 1.0
       p = nc / bb * p
       q = nc / (bb + 0.5) * q
       ps = p * s + q * r
       nc_derivative = nc_derivative + ps
       s = x / aa * s # (1 - I(x, aa+1, b)) - (1 - I(x, aa, b))
       r = x / (aa + 0.5) * r # (1 - I(x, aa+3/2, b)) - (1 - I(x, aa+1/2, b))
       term1 = p * w
       term2 = q * V
       result1 = result1 + term1
       result2 = result2 + term2
   end while ((((term1 + term2) > 0.000000000000001 * (result1 + result2)) && (s > 1E-16 * w)) || (ps > 1E-16 * nc_derivative))
   p = comppoisson(bb, nc, nc - bb)
   bb = bb + 0.5
   q = comppoisson(bb, nc, nc - bb)
   nc_derivative = x * nc_derivative * vScale
   save_result1 = result1 * vScale + p * w
   save_result2 = result2 * vScale + q * V
   ps = 1.0
   nc_dtemp = 0.0
   aa = n + a
   bb = n
   p = ptnc
   q = qtnc
   vScale = max(ptx, qtx)
   if ptx = vScale
      s = 1.0
   else
      s = ptx / vScale
   end
   if qtx = vScale
      r = 1.0
   else
      r = qtx / vScale
   end
   w = p
   V = q
   term1 = 1.0
   term2 = 1.0
   result1 = 0.0
   result2 = 0.0
   while bb > 0.0 && ((((term1 + term2) > 0.000000000000001 * (result1 + result2)) && (p > 1E-16 * w)) || (ps > 1E-16 * nc_dtemp))
      r = (aa + 0.5) / x * r
      s = aa / x * s
      ps = p * s + q * r
      nc_dtemp = nc_dtemp + ps
      p = bb / nc * p
      q = (bb + 0.5) / nc * q
      aa = aa - 1.0
      bb = bb - 1.0
      if bb = 0.0 aa = a
      r = r / (aa + b + 0.5) # (1 - I(x, aa+3/2, b)) - (1 - I(x, aa+1/2, b))
      s = s / (aa + b) # (1 - I(x, aa + 1, b)) - (1 - I(x, aa, b))
      term1 = s * w
      term2 = r * V
      result1 = result1 + term1
      result2 = result2 + term2
      w = w + p
      V = V + q
   end
   nc_dtemp = (x * nc_dtemp + p * aa * s + q * (aa + 0.5) * r) * vScale
   if x > y
      r = beta(y, b, a + (bb + 0.5))
      s = beta(y, b, a + bb)
   else
      r = compbeta(x, a + (bb + 0.5), b)
      s = compbeta(x, a + bb, b)
   end
   if b <= 1.0 vScale = vScale * b
   result1 = save_result1 + result1 * vScale + s * w
   result2 = save_result2 + result2 * vScale + r * V
   if t > 0.0
      comp_t_nc1 = 0.5 * (result1 + result2)
      nc_derivative = nc_derivative + nc_dtemp
   else
      comp_t_nc1 = 1.0 - 0.5 * (result1 - result2)
   end
end

function  inv_t_nc1(prob::Float64, df::Float64, nc::Float64, ByRef oneMinusP::Float64)::Float64
#Uses approximations in A&S 26.6.26 and 26.7.10 for to get initial estimate, the modified NR to improve it.
Dim x::Float64, y::Float64, pr::Float64, dif::Float64, temp::Float64, nc_beta_param::Float64
Dim hix::Float64, lox::Float64, test::Float64, nc_derivative::Float64
   if (prob > 0.5)
      inv_t_nc1 = comp_inv_t_nc1(1.0 - prob, df, nc, oneMinusP)
      Exit Function
   end
   nc_beta_param = nc ^ 2 / 2.0
   lox = 0.0
   hix = t_nc_limit * abs2(df)
   pr = exp(-nc_beta_param)
   if pr > prob
      if 2.0 * prob > pr
         x = invcompbeta(0.5, df / 2.0, (pr - prob) / pr, oneMinusP)
      else
         x = invbeta(0.5, df / 2.0, prob / pr, oneMinusP)
      end
      if x = 0.0
         inv_t_nc1 = 0.0
         Exit Function
      else
         temp = oneMinusP
         y = invbeta((0.5 + nc_beta_param) ^ 2 / (0.5 + 2.0 * nc_beta_param), df / 2.0, prob, oneMinusP)
         oneMinusP = (0.5 + nc_beta_param) * oneMinusP / (0.5 + nc_beta_param * (1.0 + y))
         if temp > oneMinusP
            oneMinusP = temp
         else
            x = (0.5 + 2.0 * nc_beta_param) * y / (0.5 + nc_beta_param * (1.0 + y))
         end
         if oneMinusP < cSmall
            pr = t_nc1(hix, df, nc, nc_derivative)
            if pr < prob
               inv_t_nc1 = [#VALUE!]
               oneMinusP = 0.0
               Exit Function
            end
            oneMinusP = 4.0 * cSmall
         end
      end
   else
      y = invbeta((0.5 + nc_beta_param) ^ 2 / (0.5 + 2.0 * nc_beta_param), df / 2.0, prob, oneMinusP)
      x = (0.5 + 2.0 * nc_beta_param) * y / (0.5 + nc_beta_param * (1 + y))
      oneMinusP = (0.5 + nc_beta_param) * oneMinusP / (0.5 + nc_beta_param * (1.0 + y))
      if oneMinusP < cSmall
         pr = t_nc1(hix, df, nc, nc_derivative)
         if pr < prob
            inv_t_nc1 = [#VALUE!]
            oneMinusP = 0.0
            Exit Function
         end
         oneMinusP = 4.0 * cSmall
      end
   end
   test = abs2(df * x) / abs2(oneMinusP)
   Do
      pr = t_nc1(test, df, nc, nc_derivative)
      if pr < prob
         lox = test
      else
         hix = test
      end
      if nc_derivative = 0.0
         if pr < prob
            dif = (hix - lox) / 2.0
         else
            dif = (lox - hix) / 2.0
         end
      else
         dif = -(pr * test / nc_derivative) * logdif(pr, prob)
         if df < 2.0 dif = 2.0 * dif / df
         if test + dif < lox
            if lox = 0
               dif = (lox - test) * 0.9999999999
            else
               dif = (lox - test) * 0.9
            end
         elseif test + dif > hix
            dif = (hix - test) * 0.9
         end
      end
      test = test + dif
   end while ((abs(pr - prob) > prob * 0.00000000000001) && (abs(dif) > test * 0.0000000001))
   inv_t_nc1 = test
end

function  comp_inv_t_nc1(prob::Float64, df::Float64, nc::Float64, ByRef oneMinusP::Float64)::Float64
#Uses approximations in A&S 26.6.26 and 26.7.10 for to get initial estimate, the modified NR to improve it.
Dim x::Float64, y::Float64, pr::Float64, dif::Float64, temp::Float64, nc_beta_param::Float64
Dim hix::Float64, lox::Float64, test::Float64, nc_derivative::Float64
   if (prob > 0.5)
      comp_inv_t_nc1 = inv_t_nc1(1.0 - prob, df, nc, oneMinusP)
      Exit Function
   end
   nc_beta_param = nc ^ 2 / 2.0
   lox = 0.0
   hix = t_nc_limit * abs2(df)
   pr = exp(-nc_beta_param)
   if pr > prob
      if 2.0 * prob > pr
         x = invbeta(0.5, df / 2.0, (pr - prob) / pr, oneMinusP)
      else
         x = invcompbeta(0.5, df / 2.0, prob / pr, oneMinusP)
      end
      if oneMinusP < cSmall
         pr = comp_t_nc1(hix, df, nc, nc_derivative)
         if pr > prob
            comp_inv_t_nc1 = [#VALUE!]
            oneMinusP = 0.0
            Exit Function
         end
         oneMinusP = 4.0 * cSmall
      else
         temp = oneMinusP
         y = invcompbeta((0.5 + nc_beta_param) ^ 2 / (0.5 + 2.0 * nc_beta_param), df / 2.0, prob, oneMinusP)
         oneMinusP = (0.5 + nc_beta_param) * oneMinusP / (0.5 + nc_beta_param * (1.0 + y))
         if temp < oneMinusP
            oneMinusP = temp
         else
            x = (0.5 + 2.0 * nc_beta_param) * y / (0.5 + nc_beta_param * (1.0 + y))
         end
         if oneMinusP < cSmall
            pr = comp_t_nc1(hix, df, nc, nc_derivative)
            if pr > prob
               comp_inv_t_nc1 = [#VALUE!]
               oneMinusP = 0.0
               Exit Function
            end
            oneMinusP = 4.0 * cSmall
         end
      end
   else
      y = invcompbeta((0.5 + nc_beta_param) ^ 2 / (0.5 + 2.0 * nc_beta_param), df / 2.0, prob, oneMinusP)
      x = (0.5 + 2.0 * nc_beta_param) * y / (0.5 + nc_beta_param * (1.0 + y))
      oneMinusP = (0.5 + nc_beta_param) * oneMinusP / (0.5 + nc_beta_param * (1.0 + y))
      if oneMinusP < cSmall
         pr = comp_t_nc1(hix, df, nc, nc_derivative)
         if pr > prob
            comp_inv_t_nc1 = [#VALUE!]
            oneMinusP = 0.0
            Exit Function
         end
         oneMinusP = 4.0 * cSmall
      end
   end
   test = abs2(df * x) / abs2(oneMinusP)
   dif = test
   Do
      pr = comp_t_nc1(test, df, nc, nc_derivative)
      if pr < prob
         hix = test
      else
         lox = test
      end
      if nc_derivative = 0.0
         if pr < prob
            dif = (lox - hix) / 2.0
         else
            dif = (hix - lox) / 2.0
         end
      else
         dif = (pr * test / nc_derivative) * logdif(pr, prob)
         if df < 2.0 dif = 2.0 * dif / df
         if test + dif < lox
            dif = (lox - test) * 0.9
         elseif test + dif > hix
            dif = (hix - test) * 0.9
         end
      end
      test = test + dif
   end while ((abs(pr - prob) > prob * 0.00000000000001) && (abs(dif) > test * 0.0000000001))
   comp_inv_t_nc1 = test
end

function  ncp_t_nc1(prob::Float64, t::Float64, df::Float64)::Float64
#Uses Normal approx for non-central t (A&S 26.7.10) to get initial estimate the modified NR to improve it.
Dim ncp::Float64, pr::Float64, dif::Float64, temp::Float64, deriv::Float64, checked_tnc_limit::Bool, checked_0_limit::Bool
Dim hi::Float64, lo::Float64, tnc_limit::Float64, x::Float64, y::Float64
   if (prob > 0.5)
      ncp_t_nc1 = comp_ncp_t_nc1(1.0 - prob, t, df)
      Exit Function
   end

   lo = 0.0
   tnc_limit = abs2(2.0 * nc_limit)
   hi = tnc_limit
   checked_0_limit = false
   checked_tnc_limit = false
   if t >= min(1.0, df)
      y = df / t
      x = t + y
      y = y / x
      x = t / x
   else
      x = t * t
      y = df + x
      x = x / y
      y = df / y
   end
   temp = -inv_normal(prob)
   if t > df
        ncp = t * (1.0 - 0.25 / df) + temp * abs2(t) * abs2((1.0 / t + 0.5 * t / df))
   else
        ncp = t * (1.0 - 0.25 / df) + temp * abs2((1.0 + (0.5 * t / df) * t))
   end
   ncp = max(temp, ncp)
   #Debug.Print "ncp_estimate1", ncp
   if x > 1E-200 #I think we can put more accurate bounds on when this will not deliver a sensible answer
      temp = invcompgamma(0.5 * x * df, prob) / y - 0.5
      if temp > 0
         temp = abs2(2.0 * temp)
         if temp > ncp
            ncp = temp
         end
      end
   end
   #Debug.Print "ncp_estimate2", ncp
   ncp = min(ncp, tnc_limit)
   if ncp = tnc_limit
      pr = cdf_t_nc(t, df, ncp)
      if pr > prob
         ncp_t_nc1 = [#VALUE!]
         Exit Function
      else
         checked_tnc_limit = true
      end
   end
   dif = ncp
   Do
      pr = cdf_t_nc(t, df, ncp)
      #Debug.Print ncp, pr, prob
      if ncp > 1
         deriv = cdf_t_nc(t, df, ncp * (1 - 0.000001))
         deriv = 1000000.0 * (deriv - pr) / ncp
      elseif ncp > 0.000001
         deriv = cdf_t_nc(t, df, ncp + 0.000001)
         deriv = 1000000.0 * (pr - deriv)
      elseif x < y
         deriv = comp_cdf_beta(x, 1, df / 2) * OneOverSqrTwoPi
      else
         deriv = cdf_beta(y, df / 2, 1) * OneOverSqrTwoPi
      end
      if pr < 3E-308 && deriv = 0.0
         hi = ncp
         dif = dif / 2.0
         ncp = ncp - dif
      elseif deriv = 0.0
         lo = ncp
         dif = dif / 2.0
         ncp = ncp - dif
      else
         if pr < prob
            hi = ncp
         else
            lo = ncp
         end
         dif = (pr / deriv) * logdif(pr, prob)
         if ncp + dif < lo
            dif = (lo - ncp) / 2.0
            if Not checked_0_limit && (lo = 0.0)
               temp = cdf_t_nc(t, df, lo)
               if temp < prob
                  if invtdist(prob, df) <= t
                     ncp_t_nc1 = 0.0
                  else
                     ncp_t_nc1 = [#VALUE!]
                  end
                  Exit Function
               else
                  checked_0_limit = true
               end
               dif = dif * 1.99999999
            end
         elseif ncp + dif > hi
            dif = (hi - ncp) / 2.0
            if Not checked_tnc_limit && (hi = tnc_limit)
               temp = cdf_t_nc(t, df, hi)
               if temp > prob
                  ncp_t_nc1 = [#VALUE!]
                  Exit Function
               else
                  checked_tnc_limit = true
               end
               dif = dif * 1.99999999
            end
         end
         ncp = ncp + dif
      end
   end while ((abs(pr - prob) > prob * 0.00000000000001) && (abs(dif) > abs(ncp) * 0.0000000001))
   ncp_t_nc1 = ncp
   #Debug.Print "ncp_t_nc1", ncp_t_nc1
end

function  comp_ncp_t_nc1(prob::Float64, t::Float64, df::Float64)::Float64
#Uses Normal approx for non-central t (A&S 26.7.10) to get initial estimate the modified NR to improve it.
Dim ncp::Float64, pr::Float64, dif::Float64, temp::Float64, temp1::Float64, temp2::Float64, deriv::Float64, checked_tnc_limit::Bool, checked_0_limit::Bool
Dim hi::Float64, lo::Float64, tnc_limit::Float64, x::Float64, y::Float64
   if (prob > 0.5)
      comp_ncp_t_nc1 = ncp_t_nc1(1.0 - prob, t, df)
      Exit Function
   end

   lo = 0.0
   tnc_limit = abs2(2.0 * nc_limit)
   hi = tnc_limit
   checked_0_limit = false
   checked_tnc_limit = false
   if t >= min(1.0, df)
      y = df / t
      x = t + y
      y = y / x
      x = t / x
   else
      x = t * t
      y = df + x
      x = x / y
      y = df / y
   end
   temp = -inv_normal(prob)
   temp1 = t * (1.0 - 0.25 / df)
   if t > df
        temp2 = temp * abs2(t) * abs2((1.0 / t + 0.5 * t / df))
   else
        temp2 = temp * abs2((1.0 + (0.5 * t / df) * t))
   end
   ncp = max(temp, temp1 + temp2)
   #Debug.Print "comp_ncp ncp estimate1", ncp
   if x > 1E-200 #I think we can put more accurate bounds on when this will not deliver a sensible answer
      temp = invcompgamma(0.5 * x * df, prob) / y - 0.5
      if temp > 0
         temp = abs2(2.0 * temp)
         if temp > ncp
            temp = invgamma(0.5 * x * df, prob) / y - 0.5
            if temp > 0
               ncp = abs2(2.0 * temp)
            else
               ncp = 0
            end
         else
            ncp = temp1 - temp2
         end
      else
         ncp = temp1 - temp2
      end
   else
      ncp = temp1 - temp2
   end
   ncp = min(max(0.0, ncp), tnc_limit)
   if ncp = 0.0
      pr = comp_cdf_t_nc(t, df, 0.0)
      if pr > prob
         if -invtdist(prob, df) <= t
            comp_ncp_t_nc1 = 0.0
         else
            comp_ncp_t_nc1 = [#VALUE!]
         end
         Exit Function
      elseif abs(pr - prob) <= -prob * 0.00000000000001 * log(pr)
         comp_ncp_t_nc1 = 0.0
         Exit Function
      else
         checked_0_limit = true
      end
      if x < y
         deriv = -comp_cdf_beta(x, 1, 0.5 * df) * OneOverSqrTwoPi
      else
         deriv = -cdf_beta(y, 0.5 * df, 1) * OneOverSqrTwoPi
      end
      if deriv = 0.0
         ncp = tnc_limit
      else
         ncp = (pr - prob) / deriv
         if ncp >= tnc_limit
            ncp = (pr / deriv) * logdif(pr, prob) #if these two are miles apart then best to take invgamma estimate if > 0
         end
      end
   end
   ncp = min(ncp, tnc_limit)
   if ncp = tnc_limit
      pr = comp_cdf_t_nc(t, df, ncp)
      if pr < prob
         comp_ncp_t_nc1 = [#VALUE!]
         Exit Function
      else
         checked_tnc_limit = true
      end
   end
   dif = ncp
   Do
      pr = comp_cdf_t_nc(t, df, ncp)
      #Debug.Print ncp, pr, prob
      if ncp > 1
         deriv = comp_cdf_t_nc(t, df, ncp * (1 - 0.000001))
         deriv = 1000000.0 * (pr - deriv) / ncp
      elseif ncp > 0.000001
         deriv = comp_cdf_t_nc(t, df, ncp + 0.000001)
         deriv = 1000000.0 * (deriv - pr)
      elseif x < y
         deriv = comp_cdf_beta(x, 1, 0.5 * df) * OneOverSqrTwoPi
      else
         deriv = cdf_beta(y, 0.5 * df, 1) * OneOverSqrTwoPi
      end
      if pr < 3E-308 && deriv = 0.0
         lo = ncp
         dif = dif / 2.0
         ncp = ncp - dif
      elseif deriv = 0.0
         hi = ncp
         dif = dif / 2.0
         ncp = ncp - dif
      else
         if pr > prob
            hi = ncp
         else
            lo = ncp
         end
         dif = -(pr / deriv) * logdif(pr, prob)
         if ncp + dif < lo
            dif = (lo - ncp) / 2.0
            if Not checked_0_limit && (lo = 0.0)
               temp = comp_cdf_t_nc(t, df, lo)
               if temp > prob
                  if -invtdist(prob, df) <= t
                     comp_ncp_t_nc1 = 0.0
                  else
                     comp_ncp_t_nc1 = [#VALUE!]
                  end
                  Exit Function
               else
                  checked_0_limit = true
               end
               dif = dif * 1.99999999
            end
         elseif ncp + dif > hi
            dif = (hi - ncp) / 2.0
            if Not checked_tnc_limit && (hi = tnc_limit)
               temp = comp_cdf_t_nc(t, df, hi)
               if temp < prob
                  comp_ncp_t_nc1 = [#VALUE!]
                  Exit Function
               else
                  checked_tnc_limit = true
               end
               dif = dif * 1.99999999
            end
         end
         ncp = ncp + dif
      end
   end while ((abs(pr - prob) > prob * 0.00000000000001) && (abs(dif) > abs(ncp) * 0.0000000001))
   comp_ncp_t_nc1 = ncp
   #Debug.Print "comp_ncp_t_nc1", comp_ncp_t_nc1
end

function pdf_t_nc(x::Float64, df::Float64, nc_param::Float64)::Float64
#// Calculate pdf of noncentral t
#// Deliberately set not to calculate when x and nc_param have opposite signs as the algorithm used is prone to cancellation error in these circumstances.
#// The user can access t_nc1,comp_t_nc1 directly and check on the accuracy of the results, if required
  Dim nc_derivative::Float64
  df = AlterForIntegralChecks_df(df)
  if (x < 0.0) && (nc_param <= 0.0)
     pdf_t_nc = pdf_t_nc(-x, df, -nc_param)
  elseif (df <= 0.0) || (nc_param < 0.0) || (nc_param > abs2(2.0 * nc_limit))
     pdf_t_nc = [#VALUE!]
  elseif (x < 0.0)
     pdf_t_nc = [#VALUE!]
  elseif (x = 0.0 || nc_param = 0.0)
     pdf_t_nc = exp(-nc_param ^ 2 / 2) * pdftdist(x, df)
  else
     if (df < 1.0 || x < 1.0 || x <= nc_param)
        pdf_t_nc = t_nc1(x, df, nc_param, nc_derivative)
     else
        pdf_t_nc = comp_t_nc1(x, df, nc_param, nc_derivative)
     end
     if nc_derivative < cSmall
        pdf_t_nc = exp(-nc_param ^ 2 / 2) * pdftdist(x, df)
     elseif df > 2.0
        pdf_t_nc = nc_derivative / x
     else
        pdf_t_nc = nc_derivative * (df / (2.0 * x))
     end
  end
  pdf_t_nc = GetRidOfMinusZeroes(pdf_t_nc)
end

function cdf_t_nc(x::Float64, df::Float64, nc_param::Float64)::Float64
#// Calculate cdf of noncentral t
#// Deliberately set not to calculate when x and nc_param have opposite signs as the algorithm used is prone to cancellation error in these circumstances.
#// The user can access t_nc1,comp_t_nc1 directly and check on the accuracy of the results, if required
  Dim tdistDensity::Float64, nc_derivative::Float64
  df = AlterForIntegralChecks_df(df)
  if (nc_param = 0.0)
     cdf_t_nc = tdist(x, df, tdistDensity)
  elseif (x <= 0.0) && (nc_param < 0.0)
     cdf_t_nc = comp_cdf_t_nc(-x, df, -nc_param)
  elseif (df <= 0.0) || (nc_param < 0.0) || (nc_param > abs2(2.0 * nc_limit))
     cdf_t_nc = [#VALUE!]
  elseif (x < 0.0)
     cdf_t_nc = [#VALUE!]
  elseif (df < 1.0 || x < 1.0 || x <= nc_param)
     cdf_t_nc = t_nc1(x, df, nc_param, nc_derivative)
  else
     cdf_t_nc = 1.0 - comp_t_nc1(x, df, nc_param, nc_derivative)
  end
  cdf_t_nc = GetRidOfMinusZeroes(cdf_t_nc)
end

function comp_cdf_t_nc(x::Float64, df::Float64, nc_param::Float64)::Float64
#// Calculate 1-cdf of noncentral t
#// Deliberately set not to calculate when x and nc_param have opposite signs as the algorithm used is prone to cancellation error in these circumstances.
#// The user can access t_nc1,comp_t_nc1 directly and check on the accuracy of the results, if required
  Dim tdistDensity::Float64, nc_derivative::Float64
  df = AlterForIntegralChecks_df(df)
  if (nc_param = 0.0)
     comp_cdf_t_nc = tdist(-x, df, tdistDensity)
  elseif (x <= 0.0) && (nc_param < 0.0)
     comp_cdf_t_nc = cdf_t_nc(-x, df, -nc_param)
  elseif (df <= 0.0) || (nc_param < 0.0) || (nc_param > abs2(2.0 * nc_limit))
     comp_cdf_t_nc = [#VALUE!]
  elseif (x < 0.0)
     comp_cdf_t_nc = [#VALUE!]
  elseif (df < 1.0 || x < 1.0 || x >= nc_param)
     comp_cdf_t_nc = comp_t_nc1(x, df, nc_param, nc_derivative)
  else
     comp_cdf_t_nc = 1.0 - t_nc1(x, df, nc_param, nc_derivative)
  end
  comp_cdf_t_nc = GetRidOfMinusZeroes(comp_cdf_t_nc)
end

function inv_t_nc(prob::Float64, df::Float64, nc_param::Float64)::Float64
  df = AlterForIntegralChecks_df(df)
  if (nc_param = 0.0)
     inv_t_nc = invtdist(prob, df)
  elseif (nc_param < 0.0)
     inv_t_nc = -comp_inv_t_nc(prob, df, -nc_param)
  elseif (df <= 0.0 || nc_param > abs2(2.0 * nc_limit) || prob <= 0.0 || prob >= 1.0)
     inv_t_nc = [#VALUE!]
  elseif (invcnormal(prob) < -nc_param)
     inv_t_nc = [#VALUE!]
  else
     Dim oneMinusP::Float64
     inv_t_nc = inv_t_nc1(prob, df, nc_param, oneMinusP)
  end
  inv_t_nc = GetRidOfMinusZeroes(inv_t_nc)
end

function comp_inv_t_nc(prob::Float64, df::Float64, nc_param::Float64)::Float64
  df = AlterForIntegralChecks_df(df)
  if (nc_param = 0.0)
     comp_inv_t_nc = -invtdist(prob, df)
  elseif (nc_param < 0.0)
     comp_inv_t_nc = -inv_t_nc(prob, df, -nc_param)
  elseif (df <= 0.0 || nc_param > abs2(2.0 * nc_limit) || prob <= 0.0 || prob >= 1.0)
     comp_inv_t_nc = [#VALUE!]
  elseif (invcnormal(prob) > nc_param)
     comp_inv_t_nc = [#VALUE!]
  else
     Dim oneMinusP::Float64
     comp_inv_t_nc = comp_inv_t_nc1(prob, df, nc_param, oneMinusP)
  end
  comp_inv_t_nc = GetRidOfMinusZeroes(comp_inv_t_nc)
end

function ncp_t_nc(prob::Float64, x::Float64, df::Float64)::Float64
  df = AlterForIntegralChecks_df(df)
  if (x = 0.0 && prob > 0.5)
     ncp_t_nc = -invcnormal(prob)
  elseif (x < 0)
     ncp_t_nc = -comp_ncp_t_nc(prob, -x, df)
  elseif (df <= 0.0 || prob <= 0.0 || prob >= 1.0)
     ncp_t_nc = [#VALUE!]
  else
     ncp_t_nc = ncp_t_nc1(prob, x, df)
  end
  ncp_t_nc = GetRidOfMinusZeroes(ncp_t_nc)
end

function comp_ncp_t_nc(prob::Float64, x::Float64, df::Float64)::Float64
  df = AlterForIntegralChecks_df(df)
  if (x = 0.0)
     comp_ncp_t_nc = invcnormal(prob)
  elseif (x < 0)
     comp_ncp_t_nc = -ncp_t_nc(prob, -x, df)
  elseif (df <= 0.0 || prob <= 0.0 || prob >= 1.0)
     comp_ncp_t_nc = [#VALUE!]
  else
     comp_ncp_t_nc = comp_ncp_t_nc1(prob, x, df)
  end
  comp_ncp_t_nc = GetRidOfMinusZeroes(comp_ncp_t_nc)
end

function pmf_GammaPoisson(i::Float64, gamma_shape::Float64, gamma_scale::Float64)::Float64
   Dim p::Float64, q::Float64, dfm::Float64
   q = gamma_scale / (1.0 + gamma_scale)
   p = 1.0 / (1.0 + gamma_scale)
   i = AlterForIntegralChecks_Others(i)
   if (gamma_shape <= 0.0 || gamma_scale <= 0.0)
      pmf_GammaPoisson = [#VALUE!]
   elseif (i < 0.0)
      pmf_GammaPoisson = 0
   else
      if p < q
         dfm = gamma_shape - (gamma_shape + i) * p
      else
         dfm = (gamma_shape + i) * q - i
      end
      pmf_GammaPoisson = (gamma_shape / (gamma_shape + i)) * binomialTerm(i, gamma_shape, q, p, dfm, 0.0)
   end
   pmf_GammaPoisson = GetRidOfMinusZeroes(pmf_GammaPoisson)
end

function cdf_GammaPoisson(i::Float64, gamma_shape::Float64, gamma_scale::Float64)::Float64
   Dim p::Float64, q::Float64
   q = gamma_scale / (1.0 + gamma_scale)
   p = 1.0 / (1.0 + gamma_scale)
   i = Int(i)
   if (gamma_shape <= 0.0 || gamma_scale <= 0.0)
      cdf_GammaPoisson = [#VALUE!]
   elseif (i < 0.0)
      cdf_GammaPoisson = 0.0
   elseif (p <= q)
      cdf_GammaPoisson = beta(p, gamma_shape, i + 1.0)
   else
      cdf_GammaPoisson = compbeta(q, i + 1.0, gamma_shape)
   end
   cdf_GammaPoisson = GetRidOfMinusZeroes(cdf_GammaPoisson)
end

function comp_cdf_GammaPoisson(i::Float64, gamma_shape::Float64, gamma_scale::Float64)::Float64
   Dim p::Float64, q::Float64
   q = gamma_scale / (1.0 + gamma_scale)
   p = 1.0 / (1.0 + gamma_scale)
   i = Int(i)
   if (gamma_shape <= 0.0 || gamma_scale <= 0.0)
      comp_cdf_GammaPoisson = [#VALUE!]
   elseif (i < 0.0)
      comp_cdf_GammaPoisson = 1.0
   elseif (p <= q)
      comp_cdf_GammaPoisson = compbeta(p, gamma_shape, i + 1.0)
   else
      comp_cdf_GammaPoisson = beta(q, i + 1.0, gamma_shape)
   end
   comp_cdf_GammaPoisson = GetRidOfMinusZeroes(comp_cdf_GammaPoisson)
end

function crit_GammaPoisson(gamma_shape::Float64, gamma_scale::Float64, crit_prob::Float64)::Float64
   Dim p::Float64, q::Float64
   q = gamma_scale / (1.0 + gamma_scale)
   p = 1.0 / (1.0 + gamma_scale)
   if (gamma_shape < 0.0 || gamma_scale < 0.0)
      crit_GammaPoisson = [#VALUE!]
   elseif (crit_prob < 0.0 || crit_prob >= 1.0)
      crit_GammaPoisson = [#VALUE!]
   elseif (crit_prob = 0.0)
      crit_GammaPoisson = [#VALUE!]
   else
      Dim i::Float64, pr::Float64
      crit_GammaPoisson = critnegbinom(gamma_shape, p, q, crit_prob)
      i = crit_GammaPoisson
      if p <= q
         pr = beta(p, gamma_shape, i + 1.0)
      else
         pr = compbeta(q, i + 1.0, gamma_shape)
      end
      if (pr = crit_prob)
      elseif (pr > crit_prob)
         i = i - 1.0
         if p <= q
            pr = beta(p, gamma_shape, i + 1.0)
         else
            pr = compbeta(q, i + 1.0, gamma_shape)
         end
         if (pr >= crit_prob)
            crit_GammaPoisson = i
         end
      else
         crit_GammaPoisson = i + 1.0
      end
   end
   crit_GammaPoisson = GetRidOfMinusZeroes(crit_GammaPoisson)
end

function comp_crit_GammaPoisson(gamma_shape::Float64, gamma_scale::Float64, crit_prob::Float64)::Float64
   Dim p::Float64, q::Float64
   q = gamma_scale / (1.0 + gamma_scale)
   p = 1.0 / (1.0 + gamma_scale)
   if (gamma_shape < 0.0 || gamma_scale < 0.0)
      comp_crit_GammaPoisson = [#VALUE!]
   elseif (crit_prob <= 0.0 || crit_prob > 1.0)
      comp_crit_GammaPoisson = [#VALUE!]
   elseif (crit_prob = 1.0)
      comp_crit_GammaPoisson = [#VALUE!]
   else
      Dim i::Float64, pr::Float64
      comp_crit_GammaPoisson = critcompnegbinom(gamma_shape, p, q, crit_prob)
      i = comp_crit_GammaPoisson
      if p <= q
         pr = compbeta(p, gamma_shape, i + 1.0)
      else
         pr = beta(q, i + 1.0, gamma_shape)
      end
      if (pr = crit_prob)
      elseif (pr < crit_prob)
         i = i - 1.0
         if p <= q
            pr = compbeta(p, gamma_shape, i + 1.0)
         else
            pr = beta(q, i + 1.0, gamma_shape)
         end
         if (pr <= crit_prob)
            comp_crit_GammaPoisson = i
         end
      else
         comp_crit_GammaPoisson = i + 1.0
      end
   end
   comp_crit_GammaPoisson = GetRidOfMinusZeroes(comp_crit_GammaPoisson)
end

function  PBB(i::Float64, ssmi::Float64, beta_shape1::Float64, beta_shape2::Float64)::Float64
    hTerm = hypergeometricTerm(i, ssmi, beta_shape1, beta_shape2)
    PBB = (beta_shape1 / (i + beta_shape1)) * (beta_shape2 / (beta_shape1 + beta_shape2)) * ((i + ssmi + beta_shape1 + beta_shape2) / (ssmi + beta_shape2)) * hTerm
end

function  PBNB(i::Float64, r::Float64, beta_shape1::Float64, beta_shape2::Float64)::Float64
    hTerm = hypergeometricTerm(i, r, beta_shape2, beta_shape1)
    PBNB = (beta_shape2 / (beta_shape1 + beta_shape2)) * (r / (beta_shape1 + r)) * beta_shape1 * (i + beta_shape1 + r + beta_shape2) / ((i + r) * (i + beta_shape2)) * hTerm
end

function pmf_BetaNegativeBinomial(i::Float64, r::Float64, beta_shape1::Float64, beta_shape2::Float64)::Float64
   i = AlterForIntegralChecks_Others(i)
   if (r <= 0.0 || beta_shape1 <= 0.0 || beta_shape2 <= 0.0)
      pmf_BetaNegativeBinomial = [#VALUE!]
   elseif i < 0
      pmf_BetaNegativeBinomial = 0.0
   else
      pmf_BetaNegativeBinomial = (beta_shape2 / (beta_shape1 + beta_shape2)) * (r / (beta_shape1 + r)) * beta_shape1 * (i + beta_shape1 + r + beta_shape2) / ((i + r) * (i + beta_shape2)) * hypergeometricTerm(i, r, beta_shape2, beta_shape1)
   end
   pmf_BetaNegativeBinomial = GetRidOfMinusZeroes(pmf_BetaNegativeBinomial)
end

function  CBNB0(i::Float64, r::Float64, beta_shape1::Float64, beta_shape2::Float64, toBeAdded::Float64)::Float64
   Dim ha1::Float64, hprob::Float64, hswap::Bool
   Dim mrb2::Float64, other::Float64, temp::Float64
   if (r < 2.0 || beta_shape2 < 2.0)
#One assumption here that i is integral or greater than 4.
      mrb2 = max(r, beta_shape2)
      other = min(r, beta_shape2)
      CBNB0 = PBB(i, other, mrb2, beta_shape1)
      if i = 0.0 Exit Function
      CBNB0 = CBNB0 * (1.0 + i * (other + beta_shape1) / (((i - 1.0) + mrb2) * (other + 1.0)))
      if i = 1.0 Exit Function
      i = i - 2.0
      other = other + 2.0
      temp = PBB(i, mrb2, other, beta_shape1)
      if i = 0.0
         CBNB0 = CBNB0 + temp
         Exit Function
      end
      CBNB0 = CBNB0 + temp * (1.0 + i * (mrb2 + beta_shape1) / (((i - 1.0) + other) * (mrb2 + 1.0)))
      if i = 1.0 Exit Function
      i = i - 2.0
      mrb2 = mrb2 + 2.0
      CBNB0 = CBNB0 + CBNB0(i, mrb2, beta_shape1, other, CBNB0)
   elseif (beta_shape1 < 1.0)
      mrb2 = max(r, beta_shape2)
      other = min(r, beta_shape2)
      CBNB0 = hypergeometric(i, mrb2 - 1.0, other, beta_shape1, false, ha1, hprob, hswap)
      if hswap
         temp = PBB(mrb2 - 1.0, beta_shape1, i + 1.0, other)
         if (toBeAdded + (CBNB0 - temp)) < 0.01 * (toBeAdded + (CBNB0 + temp))
            CBNB0 = CBNB2(i, mrb2, beta_shape1, other)
         else
            CBNB0 = CBNB0 - temp
         end
      elseif ha1 < -0.9 * beta_shape1 / (beta_shape1 + other)
         CBNB0 = [#VALUE!]
      else
         CBNB0 = hprob * (beta_shape1 / (beta_shape1 + other) + ha1)
      end
   else
      CBNB0 = hypergeometric(i, r, beta_shape2, beta_shape1 - 1.0, false, ha1, hprob, hswap)
   end
end

function  CBNB2(i::Float64, r::Float64, beta_shape1::Float64, beta_shape2::Float64)::Float64
   Dim j::Float64, ss::Float64, bs2::Float64, temp::Float64, d1::Float64, d2::Float64, d_count::Float64, pbbval::Float64
   #In general may be a good idea to take min(i, beta_shape1) down to just above 0 and then work on max(i, beta_shape1)
   ss = min(r, beta_shape2)
   bs2 = max(r, beta_shape2)
   r = ss
   beta_shape2 = bs2
   d1 = (i + 0.5) * (beta_shape1 + 0.5) - (bs2 - 1.5) * (ss - 0.5)
   if d1 < 0.0
      CBNB2 = CBNB0(i, ss, beta_shape1, bs2, 0.0)
      Exit Function
   end
   d1 = Int(d1 / (bs2 + beta_shape1 - 1.0)) + 10.0
   if ss + d1 > bs2 d1 = Int(bs2 - ss)
   ss = ss + d1
   j = i - d1
   d2 = (j + 0.5) * (beta_shape1 + 0.5) - (bs2 - 1.5) * (ss - 0.5)
   if d2 < 0.0
      d2 = 10.0
   else
      temp = bs2 + ss + 2.0 * beta_shape1 - 1.0
      d2 = Int((abs2(temp ^ 2 + 4.0 * d2) - temp) / 2.0) + 10.0
   end
   if 2.0 * d2 > i
      d2 = Int(i / 2.0)
   end
   pbbval = PBB(i, r, beta_shape2, beta_shape1)
   ss = ss + d2
   bs2 = bs2 + d2
   j = j - 2.0 * d2
   CBNB2 = CBNB0(j, ss, beta_shape1, bs2, 0.0)
   temp = 1.0
   d_count = d2 - 2.0
   j = j + 1.0
   while d_count >= 0.0
      j = j + 1.0
      bs2 = beta_shape2 + d_count
      d_count = d_count - 1.0
      temp = 1.0 + (j * (bs2 + beta_shape1) / ((j + ss - 1.0) * (bs2 + 1.0))) * temp
   end
   j = i - d2 - d1
   temp = (ss * (j + bs2)) / (bs2 * (j + ss)) * temp
   d_count = d1 + d2 - 1.0
   while d_count >= 0
      j = j + 1.0
      ss = r + d_count
      d_count = d_count - 1.0
      temp = 1.0 + (j * (ss + beta_shape1) / ((j + bs2 - 1.0) * (ss + 1.0))) * temp
   end
   CBNB2 = CBNB2 + temp * pbbval
   Exit Function
end

function cdf_BetaNegativeBinomial(i::Float64, r::Float64, beta_shape1::Float64, beta_shape2::Float64)::Float64
   i = Int(i)
   if (r <= 0.0 || beta_shape1 <= 0.0 || beta_shape2 <= 0.0)
      cdf_BetaNegativeBinomial = [#VALUE!]
   elseif i < 0
      cdf_BetaNegativeBinomial = 0.0
   else
      cdf_BetaNegativeBinomial = CBNB0(i, r, beta_shape1, beta_shape2, 0.0)
   end
   cdf_BetaNegativeBinomial = GetRidOfMinusZeroes(cdf_BetaNegativeBinomial)
end

function comp_cdf_BetaNegativeBinomial(i::Float64, r::Float64, beta_shape1::Float64, beta_shape2::Float64)::Float64
   Dim ha1::Float64, hprob::Float64, hswap::Bool
   Dim mrb2::Float64, other::Float64, temp::Float64, mnib1::Float64, mxib1::Float64, swap::Float64, max_iterations::Float64
   i = Int(i)
   mrb2 = max(r, beta_shape2)
   other = min(r, beta_shape2)
   if (other <= 0.0 || beta_shape1 <= 0.0)
      comp_cdf_BetaNegativeBinomial = [#VALUE!]
   elseif i < 0.0
      comp_cdf_BetaNegativeBinomial = 1.0
   elseif (i = 0.0) || ((i < 1000000.0) && (other < 0.001) && (beta_shape1 > 50.0 * other) && (100.0 * i * beta_shape1 < mrb2))
      comp_cdf_BetaNegativeBinomial = ccBNB5(i, mrb2, beta_shape1, other)
   elseif (mrb2 >= 100.0 || other > 20.0 || (mrb2 >= 5.0 && (other - 0.5) * (mrb2 - 0.5) > (i + 0.5) * (beta_shape1 + 0.5)))
      comp_cdf_BetaNegativeBinomial = CBNB0(mrb2 - 1.0, i + 1.0, other, beta_shape1, 0.0)
   else
      comp_cdf_BetaNegativeBinomial = 0.0
      temp = 0.0
      i = i + 1.0
      if other >= 1.0
         mrb2 = mrb2 - 1.0
         other = other - 1.0
         temp = hypergeometricTerm(i, mrb2, other, beta_shape1)
         comp_cdf_BetaNegativeBinomial = temp
         while (other >= 1.0) && (temp > 1E-16 * comp_cdf_BetaNegativeBinomial)
            i = i + 1.0
            beta_shape1 = beta_shape1 + 1.0
            temp = temp * (mrb2 * other) / (i * beta_shape1)
            mrb2 = mrb2 - 1.0
            other = other - 1.0
            comp_cdf_BetaNegativeBinomial = comp_cdf_BetaNegativeBinomial + temp
         end
         if other >= 1.0 Exit Function
         i = i + 1.0
         beta_shape1 = beta_shape1 + 1.0
      end
      if mrb2 >= 1.0
         mxib1 = max(i, beta_shape1)
         mnib1 = min(i, beta_shape1)
         if temp = 0.0
            mrb2 = mrb2 - 1.0
            temp = PBB(mnib1, mrb2, other, mxib1)
         else #temp is hypergeometricTerm(mnib1-1, mrb2, other, mxib1-1)
            temp = temp * other * mrb2
            mrb2 = mrb2 - 1.0
            temp = temp / (mnib1 * (mrb2 + mxib1))
         end
         comp_cdf_BetaNegativeBinomial = comp_cdf_BetaNegativeBinomial + temp
         while (mrb2 >= 1.0) && (temp > 1E-16 * comp_cdf_BetaNegativeBinomial)
            temp = temp * mrb2 * (mnib1 + other)
            mnib1 = mnib1 + 1.0
            if mnib1 > mxib1
               swap = mxib1
               mxib1 = mnib1
               mnib1 = swap
            end
#Block below not required if hypergeometric block included above and therefore other guaranteed < 1 <= mrb2
            #if mrb2 < other
            #   swap = other
            #   other = mrb2
            #   mrb2 = swap
            #end
            mrb2 = mrb2 - 1.0
            temp = temp / ((mrb2 + mxib1) * mnib1)
            comp_cdf_BetaNegativeBinomial = comp_cdf_BetaNegativeBinomial + temp
         end
         if mrb2 >= 1.0 Exit Function
         temp = temp * mrb2 / (mnib1 + mrb2)
      else
         mxib1 = beta_shape1
         mnib1 = i
         if temp = 0.0
            temp = pBNB(mnib1, mrb2, mxib1, other)
         else
            temp = temp * mrb2 * other / (i * (mrb2 + other + mnib1 + mxib1 + -1))
         end
         comp_cdf_BetaNegativeBinomial = comp_cdf_BetaNegativeBinomial + temp
      end
      max_iterations = 60.0
      Do
         temp = temp * (mnib1 + mrb2) * (mnib1 + other) / (mnib1 + mxib1 + mrb2 + other)
         mnib1 = mnib1 + 1.0
         if mxib1 < mnib1
            swap = mxib1
            mxib1 = mnib1
            mnib1 = swap
         end
         temp = temp / mnib1
         comp_cdf_BetaNegativeBinomial = comp_cdf_BetaNegativeBinomial + temp
      end Until (temp <= 1E-16 * comp_cdf_BetaNegativeBinomial) || (mnib1 + mxib1 > max_iterations)
      temp = temp * (mnib1 + mrb2) * (mnib1 + other) / ((mnib1 + 1.0) * mxib1)
      comp_cdf_BetaNegativeBinomial = comp_cdf_BetaNegativeBinomial + temp
      mnib1 = mnib1 + 1.0
      mrb2 = mrb2 - 1.0
      other = other - 1.0
      Do
         mnib1 = mnib1 + 1.0
         mxib1 = mxib1 + 1.0
         temp = temp * (mrb2 * other) / (mnib1 * mxib1)
         mrb2 = mrb2 - 1.0
         other = other - 1.0
         comp_cdf_BetaNegativeBinomial = comp_cdf_BetaNegativeBinomial + temp
      end Until abs(temp) <= 1E-16 * comp_cdf_BetaNegativeBinomial
   end
   comp_cdf_BetaNegativeBinomial = GetRidOfMinusZeroes(comp_cdf_BetaNegativeBinomial)
end

function  critbetanegbinomial(a::Float64, b::Float64, r::Float64, cprob::Float64)::Float64
#//i such that Pr(betanegbinomial(i,r,a,b)) >= cprob and  Pr(betanegbinomial(i-1,r,a,b)) < cprob
   if (cprob > 0.5)
      critbetanegbinomial = critcompbetanegbinomial(a, b, r, 1.0 - cprob)
      Exit Function
   end
   Dim pr::Float64, tpr::Float64
   Dim i::Float64, temp::Float64, temp2::Float64, oneMinusP::Float64
   if b > r
      i = b
      b = r
      r = i
   end
   if (a < 10.0 || b < 10.0)
      if r < a && a < 1.0
         pr = cprob * a / r
      else
         pr = cprob
      end
      i = invcompbeta(a, b, pr, oneMinusP)
   else
      pr = r / (r + a + b - 1.0)
      i = invcompbeta(a * pr, b * pr, cprob, oneMinusP)
   end
   if i = 0.0
      i = max_crit / 2.0
   else
      i = r * (oneMinusP / i)
      if i >= max_crit
         i = max_crit - 1.0
      end
   end
   while (true)
      if (i < 0.0)
         i = 0.0
      end
      i = Int(i + 0.5)
      if (i >= max_crit)
         critbetanegbinomial = [#VALUE!]
         Exit Function
      end
      pr = CBNB0(i, r, a, b, 0.0)
      tpr = 0.0
      if (pr > cprob * (1 + cfSmall))
         if (i = 0.0)
            critbetanegbinomial = 0.0
            Exit Function
         end
         tpr = pmf_BetaNegativeBinomial(i, r, a, b)
         if (pr < (1.0 + 0.00001) * tpr)
            i = i - 1.0
            tpr = tpr * (((i + 1.0) * (i + a + r + b)) / ((i + r) * (i + b)))
            while (tpr > cprob)
               i = i - 1.0
               tpr = tpr * (((i + 1.0) * (i + a + r + b)) / ((i + r) * (i + b)))
            end
         else
            pr = pr - tpr
            if (pr < cprob)
               critbetanegbinomial = i
               Exit Function
            end
            i = i - 1.0
            if (i = 0.0)
               critbetanegbinomial = 0.0
               Exit Function
            end
            temp = (pr - cprob) / tpr
            if (temp > 10.0)
               temp = Int(temp + 0.5)
               if temp > i
                  i = i / 10.0
               else
                  i = Int(i - temp)
                  temp2 = pmf_BetaNegativeBinomial(i, r, a, b)
                  i = i - temp * (tpr - temp2) / (2.0 * temp2)
               end
            else
               tpr = tpr * (((i + 1.0) * (i + a + r + b)) / ((i + r) * (i + b)))
               pr = pr - tpr
               if (pr < cprob)
                  critbetanegbinomial = i
                  Exit Function
               end
               temp2 = (pr - cprob) / tpr
               if (temp2 < temp - 0.9)
                  while (pr >= cprob)
                     i = i - 1.0
                     tpr = tpr * (((i + 1.0) * (i + a + r + b)) / ((i + r) * (i + b)))
                     pr = pr - tpr
                  end
                  critbetanegbinomial = i
                  Exit Function
               else
                  temp = Int(log(cprob / pr) / log((((i + 1.0) * (i + a + r + b)) / ((i + r) * (i + b)))) + 0.5)
                  i = i - temp
                  temp2 = pmf_BetaNegativeBinomial(i, r, a, b)
                  if (temp2 > nearly_zero)
                     temp = log((cprob / pr) * (tpr / temp2)) / log((((i + 1.0) * (i + a + r + b)) / ((i + r) * (i + b))))
                     i = i - temp
                  end
               end
            end
         end
      elseif ((1.0 + cfSmall) * pr < cprob)
         while ((tpr < cSmall) && (pr < cprob))
            i = i + 1.0
            tpr = pmf_BetaNegativeBinomial(i, r, a, b)
            pr = pr + tpr
            if pr = 0.0 || 1E+100 * pr < cprob
               tpr = cSmall
            end
         end
         if pr > 0.0
            temp = (cprob - pr) / tpr
         else
            temp = max_crit
         end
         if temp <= 0.0
            critbetanegbinomial = i
            Exit Function
         elseif temp < 10.0
            while (pr < cprob)
               tpr = tpr * (((i + r) * (i + b)) / ((i + 1.0) * (i + a + r + b)))
               pr = pr + tpr
               i = i + 1.0
            end
            critbetanegbinomial = i
            Exit Function
         elseif i = max_crit
            critbetanegbinomial = [#VALUE!]
            Exit Function
         elseif i + temp > max_crit
            i = max_crit - 1
         else
            i = Int(i + temp)
            temp2 = pmf_BetaNegativeBinomial(i, r, a, b)
            if temp2 > 0.0 i = i + temp * (tpr - temp2) / (2.0 * temp2)
         end
      else
         critbetanegbinomial = i
         Exit Function
      end
   end
end

function  critcompbetanegbinomial(a::Float64, b::Float64, r::Float64, cprob::Float64)::Float64
#//i such that 1-Pr(betanegbinomial(i,r,a,b)) > cprob and  1-Pr(betanegbinomial(i-1,r,a,b)) <= cprob
   if (cprob > 0.5)
      critcompbetanegbinomial = critbetanegbinomial(a, b, r, 1.0 - cprob)
      Exit Function
   end
   Dim pr::Float64, tpr::Float64, i_smallest::Float64
   Dim i::Float64, temp::Float64, temp2::Float64, oneMinusP::Float64
   i_smallest = 0.0
   if b > r
      i = b
      b = r
      r = i
   end
   if (a < 10.0 || b < 10.0)
      if r < a && a < 1.0
         pr = cprob * a / r
      else
         pr = cprob
      end
      i = invbeta(a, b, pr, oneMinusP)
   else
      pr = r / (r + a + b - 1.0)
      i = invbeta(a * pr, b * pr, cprob, oneMinusP)
   end
   if i = 0.0
      i = max_crit / 2.0
   else
      i = r * (oneMinusP / i)
      if i >= max_crit
         i = max_crit - 1.0
      end
   end
   while (true)
      if (i < 0.0)
         i = 0.0
      end
      i = Int(i + 0.5)
      if (i >= max_crit)
         critcompbetanegbinomial = [#VALUE!]
         Exit Function
      end
      pr = comp_cdf_BetaNegativeBinomial(i, r, a, b)
      tpr = 0.0
      if (pr > cprob * (1 + cfSmall))
         i = i + 1.0
         i_smallest = i
         tpr = pmf_BetaNegativeBinomial(i, r, a, b)
         if (pr < (1.00001) * tpr)
            while (tpr > cprob)
               tpr = tpr * (((i + r) * (i + b)) / ((i + 1.0) * (i + a + r + b)))
               i = i + 1.0
            end
         else
            pr = pr - tpr
            if (pr <= cprob)
               critcompbetanegbinomial = i
               Exit Function
            end
            temp = (pr - cprob) / tpr
            if (temp > 10.0)
               temp = Int(temp + 0.5)
               i = i + temp
               temp2 = pmf_BetaNegativeBinomial(i, r, a, b)
               i = i + temp * (tpr - temp2) / (2.0 * temp2)
            else
               tpr = tpr * (((i + r) * (i + b)) / ((i + 1.0) * (i + a + r + b)))
               i = i + 1.0
               pr = pr - tpr
               if (pr <= cprob)
                  critcompbetanegbinomial = i
                  Exit Function
               end
               temp2 = (pr - cprob) / tpr
               if (temp2 < temp - 0.9)
                  while (pr > cprob)
                     tpr = tpr * (((i + r) * (i + b)) / ((i + 1.0) * (i + a + r + b)))
                     i = i + 1.0
                     pr = pr - tpr
                  end
                  critcompbetanegbinomial = i
                  Exit Function
               else
                  temp = Int(log(cprob / pr) / log((((i + r - 1.0) * (i + b - 1.0)) / (i * (i + a + r + b - 1.0)))) + 0.5)
                  i = i + temp
                  temp2 = pmf_BetaNegativeBinomial(i, r, a, b)
                  if (temp2 > nearly_zero)
                     temp = log((cprob / pr) * (tpr / temp2)) / log((((i + r - 1.0) * (i + b - 1.0)) / (i * (i + a + r + b - 1.0))))
                     i = i + temp
                  end
               end
            end
         end
      elseif pr < excel0
         i = (i_smallest + i) / 2
      elseif ((1.0 + cfSmall) * pr < cprob)
         while ((tpr < cSmall) && (pr <= cprob))
            tpr = pmf_BetaNegativeBinomial(i, r, a, b)
            pr = pr + tpr
            i = i - 1.0
         end
         temp = (cprob - pr) / tpr
         if temp <= 0.0
            critcompbetanegbinomial = i + 1.0
            Exit Function
         elseif temp < 100.0 || i < 1000.0
            while (pr <= cprob)
               tpr = tpr * (((i + 1.0) * (i + a + r + b)) / ((i + r) * (i + b)))
               pr = pr + tpr
               i = i - 1.0
            end
            critcompbetanegbinomial = i + 1.0
            Exit Function
         elseif temp > i
            i = i / 10.0
         else
            i = Int(i - temp)
            temp2 = pmf_BetaNegativeBinomial(i, r, a, b)
            if temp2 > 0.0 i = i - temp * (tpr - temp2) / (2.0 * temp2)
         end
      else
         critcompbetanegbinomial = i
         Exit Function
      end
   end
end

function crit_BetaNegativeBinomial(r::Float64, beta_shape1::Float64, beta_shape2::Float64, crit_prob::Float64)::Float64
   if (beta_shape1 <= 0.0 || beta_shape2 <= 0.0 || r <= 0.0)
      crit_BetaNegativeBinomial = [#VALUE!]
   elseif (crit_prob < 0.0 || crit_prob >= 1.0)
      crit_BetaNegativeBinomial = [#VALUE!]
   elseif (crit_prob = 0.0)
      crit_BetaNegativeBinomial = [#VALUE!]
   else
      Dim i::Float64, pr::Float64
      i = critbetanegbinomial(beta_shape1, beta_shape2, r, crit_prob)
      crit_BetaNegativeBinomial = i
      pr = cdf_BetaNegativeBinomial(i, r, beta_shape1, beta_shape2)
      if (pr = crit_prob)
      elseif (pr > crit_prob)
         i = i - 1.0
         pr = cdf_BetaNegativeBinomial(i, r, beta_shape1, beta_shape2)
         if (pr >= crit_prob)
            crit_BetaNegativeBinomial = i
         end
      else
         crit_BetaNegativeBinomial = i + 1.0
      end
   end
   crit_BetaNegativeBinomial = GetRidOfMinusZeroes(crit_BetaNegativeBinomial)
end

function comp_crit_BetaNegativeBinomial(r::Float64, beta_shape1::Float64, beta_shape2::Float64, crit_prob::Float64)::Float64
   if (beta_shape1 <= 0.0 || beta_shape2 <= 0.0 || r <= 0.0)
      comp_crit_BetaNegativeBinomial = [#VALUE!]
   elseif (crit_prob <= 0.0 || crit_prob > 1.0)
      comp_crit_BetaNegativeBinomial = [#VALUE!]
   elseif (crit_prob = 1.0)
      comp_crit_BetaNegativeBinomial = 0.0
   else
      Dim i::Float64, pr::Float64
      i = critcompbetanegbinomial(beta_shape1, beta_shape2, r, crit_prob)
      comp_crit_BetaNegativeBinomial = i
      pr = comp_cdf_BetaNegativeBinomial(i, r, beta_shape1, beta_shape2)
      if (pr = crit_prob)
      elseif (pr < crit_prob)
         i = i - 1.0
         pr = comp_cdf_BetaNegativeBinomial(i, r, beta_shape1, beta_shape2)
         if (pr <= crit_prob)
            comp_crit_BetaNegativeBinomial = i
         end
      else
         comp_crit_BetaNegativeBinomial = i + 1.0
      end
   end
   comp_crit_BetaNegativeBinomial = GetRidOfMinusZeroes(comp_crit_BetaNegativeBinomial)
end

function pmf_BetaBinomial(i::Float64, sample_size::Float64, beta_shape1::Float64, beta_shape2::Float64)::Float64
   i = AlterForIntegralChecks_Others(i)
   sample_size = AlterForIntegralChecks_Others(sample_size)
   if (beta_shape1 <= 0.0 || beta_shape2 <= 0.0 || sample_size < 0.0)
      pmf_BetaBinomial = [#VALUE!]
   elseif i < 0 || i > sample_size
      pmf_BetaBinomial = 0.0
   else
      pmf_BetaBinomial = (beta_shape1 / (i + beta_shape1)) * (beta_shape2 / (beta_shape1 + beta_shape2)) * ((sample_size + beta_shape1 + beta_shape2) / (sample_size - i + beta_shape2)) * hypergeometricTerm(i, sample_size - i, beta_shape1, beta_shape2)
   end
   pmf_BetaBinomial = GetRidOfMinusZeroes(pmf_BetaBinomial)
end

function cdf_BetaBinomial(i::Float64, sample_size::Float64, beta_shape1::Float64, beta_shape2::Float64)::Float64
   i = Int(i)
   sample_size = AlterForIntegralChecks_Others(sample_size)
   if (beta_shape1 <= 0.0 || beta_shape2 <= 0.0 || sample_size < 0.0)
      cdf_BetaBinomial = [#VALUE!]
   elseif i < 0.0
      cdf_BetaBinomial = 0.0
   else
      i = i + 1.0
      cdf_BetaBinomial = comp_cdf_BetaNegativeBinomial(sample_size - i, i, beta_shape1, beta_shape2)
   end
end

function comp_cdf_BetaBinomial(i::Float64, sample_size::Float64, beta_shape1::Float64, beta_shape2::Float64)::Float64
   i = Int(i)
   sample_size = AlterForIntegralChecks_Others(sample_size)
   if (beta_shape1 <= 0.0 || beta_shape2 <= 0.0 || sample_size < 0.0)
      comp_cdf_BetaBinomial = [#VALUE!]
   elseif i < 0.0
      comp_cdf_BetaBinomial = 1.0
   elseif i >= sample_size
      comp_cdf_BetaBinomial = 0.0
   else
      comp_cdf_BetaBinomial = comp_cdf_BetaNegativeBinomial(i, sample_size - i, beta_shape2, beta_shape1)
   end
end

function  critbetabinomial(a::Float64, b::Float64, ss::Float64, cprob::Float64)::Float64
#//i such that Pr(betabinomial(i,ss,a,b)) >= cprob and  Pr(betabinomial(i-1,ss,a,b)) < cprob
   if (cprob > 0.5)
      critbetabinomial = critcompbetabinomial(a, b, ss, 1.0 - cprob)
      Exit Function
   end
   Dim pr::Float64, tpr::Float64
   Dim i::Float64, temp::Float64, temp2::Float64, oneMinusP::Float64
   if (a + b < 1.0)
      i = invbeta(a, b, cprob, oneMinusP) * ss
   else
      pr = ss / (ss + a + b - 1.0)
      i = invbeta(a * pr, b * pr, cprob, oneMinusP) * ss
   end
   while (true)
      if (i < 0.0)
         i = 0.0
      elseif (i > ss)
         i = ss
      end
      i = Int(i + 0.5)
      if (i >= max_discrete)
         critbetabinomial = i
         Exit Function
      end
      pr = cdf_BetaBinomial(i, ss, a, b)
      tpr = 0.0
      if (pr >= cprob * (1 + cfSmall))
         if (i = 0.0)
            critbetabinomial = 0.0
            Exit Function
         end
         tpr = pmf_BetaBinomial(i, ss, a, b)
         pr = pr - tpr
         if (pr < cprob)
            critbetabinomial = i
            Exit Function
         end
         tpr = tpr * (i * ((ss - i) + b)) / ((a + i - 1.0) * (ss - i + 1.0))
         i = i - 1.0
         if (pr < (1.0 + 0.00001) * tpr)
            while (tpr > cprob)
               tpr = tpr * (i * ((ss - i) + b)) / ((a + i - 1.0) * (ss - i + 1.0))
               i = i - 1.0
            end
         else
            if (i = 0.0)
               critbetabinomial = 0.0
               Exit Function
            end
            temp = (pr - cprob) / tpr
            if (temp > 10)
               temp = Int(temp + 0.5)
               i = i - temp
               temp2 = pmf_BetaBinomial(i, ss, a, b)
               i = i - temp * (tpr - temp2) / (2.0 * temp2)
            else
               tpr = tpr * (i * ((ss - i) + b)) / ((a + i - 1.0) * (ss - i + 1.0))
               pr = pr - tpr
               if (pr < cprob)
                  critbetabinomial = i
                  Exit Function
               end
               i = i - 1.0
               temp2 = (pr - cprob) / tpr
               if (temp2 < temp - 0.9)
                  while (pr >= cprob)
                     tpr = tpr * (i * ((ss - i) + b)) / ((a + i - 1.0) * (ss - i + 1.0))
                     pr = pr - tpr
                     i = i - 1.0
                  end
                  critbetabinomial = i + 1.0
                  Exit Function
               else
                  temp = Int(log(cprob / pr) / log((i * ((ss - i) + b)) / ((a + i - 1.0) * (ss - i + 1.0))) + 0.5)
                  i = i - temp
                  temp2 = pmf_BetaBinomial(i, ss, a, b)
                  if (temp2 > nearly_zero)
                     temp = log((cprob / pr) * (tpr / temp2)) / log((i * ((ss - i) + b)) / ((a + i - 1.0) * (ss - i + 1.0)))
                     i = i - temp
                  end
               end
            end
         end
      elseif ((1.0 + cfSmall) * pr < cprob)
         while ((tpr < cSmall) && (pr < cprob))
            i = i + 1.0
            tpr = pmf_BetaBinomial(i, ss, a, b)
            pr = pr + tpr
         end
         temp = (cprob - pr) / tpr
         if temp <= 0.0
            critbetabinomial = i
            Exit Function
         elseif temp < 10.0
            while (pr < cprob)
               i = i + 1.0
               tpr = tpr * ((a + i - 1.0) * (ss - i + 1.0)) / (i * ((ss - i) + b))
               pr = pr + tpr
            end
            critbetabinomial = i
            Exit Function
         elseif temp > 4E+15
            i = 4E+15
         else
            i = Int(i + temp)
            temp2 = pmf_BetaBinomial(i, ss, a, b)
            if temp2 > 0.0 i = i + temp * (tpr - temp2) / (2.0 * temp2)
         end
      else
         critbetabinomial = i
         Exit Function
      end
   end
end

function  critcompbetabinomial(a::Float64, b::Float64, ss::Float64, cprob::Float64)::Float64
#//i such that 1-Pr(betabinomial(i,ss,a,b)) > cprob and  1-Pr(betabinomial(i-1,ss,a,b)) <= cprob
   if (cprob > 0.5)
      critcompbetabinomial = critbetabinomial(a, b, ss, 1.0 - cprob)
      Exit Function
   end
   Dim pr::Float64, tpr::Float64
   Dim i::Float64, temp::Float64, temp2::Float64, oneMinusP::Float64
   if (a + b < 1.0)
      i = invcompbeta(a, b, cprob, oneMinusP) * ss
   else
      pr = ss / (ss + a + b - 1.0)
      i = invcompbeta(a * pr, b * pr, cprob, oneMinusP) * ss
   end
   while (true)
      if (i < 0.0)
         i = 0.0
      elseif (i > ss)
         i = ss
      end
      i = Int(i + 0.5)
      if (i >= max_discrete)
         critcompbetabinomial = i
         Exit Function
      end
      pr = comp_cdf_BetaBinomial(i, ss, a, b)
      tpr = 0.0
      if (pr >= cprob * (1 + cfSmall))
         i = i + 1.0
         tpr = pmf_BetaBinomial(i, ss, a, b)
         if (pr < (1.00001) * tpr)
            while (tpr > cprob)
               i = i + 1.0
               temp = ss + b - i
               if temp = 0.0 Exit Do
               tpr = tpr * ((a + i - 1.0) * (ss - i + 1.0)) / (i * temp)
            end
         else
            pr = pr - tpr
            if (pr <= cprob)
               critcompbetabinomial = i
               Exit Function
            end
            temp = (pr - cprob) / tpr
            if (temp > 10.0)
               temp = Int(temp + 0.5)
               i = i + temp
               temp2 = pmf_BetaBinomial(i, ss, a, b)
               i = i + temp * (tpr - temp2) / (2.0 * temp2)
            else
               i = i + 1.0
               tpr = tpr * ((a + i - 1.0) * (ss - i + 1.0)) / (i * (ss + b - i))
               pr = pr - tpr
               if (pr <= cprob)
                  critcompbetabinomial = i
                  Exit Function
               end
               temp2 = (pr - cprob) / tpr
               if (temp2 < temp - 0.9)
                  while (pr > cprob)
                     i = i + 1.0
                     tpr = tpr * ((a + i - 1.0) * (ss - i + 1.0)) / (i * (ss + b - i))
                     pr = pr - tpr
                  end
                  critcompbetabinomial = i
                  Exit Function
               else
                  temp = Int(log(cprob / pr) / log(((a + i - 1.0) * (ss - i + 1.0)) / (i * (ss + b - i))) + 0.5)
                  i = i + temp
                  temp2 = pmf_BetaBinomial(i, ss, a, b)
                  if (temp2 > nearly_zero)
                     temp = log((cprob / pr) * (tpr / temp2)) / log(((a + i - 1.0) * (ss - i + 1.0)) / (i * (ss + b - i)))
                     i = i + temp
                  end
               end
            end
         end
      elseif ((1.0 + cfSmall) * pr < cprob)
         while ((tpr < cSmall) && (pr <= cprob))
            tpr = pmf_BetaBinomial(i, ss, a, b)
            pr = pr + tpr
            i = i - 1.0
         end
         temp = (cprob - pr) / tpr
         if temp <= 0.0
            critcompbetabinomial = i + 1.0
            Exit Function
         elseif temp < 100.0 || i < 1000.0
            while (pr <= cprob)
               tpr = tpr * ((i + 1.0) * (ss + b - i - 1.0)) / ((a + i) * (ss - i))
               pr = pr + tpr
               i = i - 1.0
            end
            critcompbetabinomial = i + 1.0
            Exit Function
         elseif temp > i
            i = i / 10.0
         else
            i = Int(i - temp)
            temp2 = pmf_BetaNegativeBinomial(i, ss, a, b)
            if temp2 > 0.0 i = i - temp * (tpr - temp2) / (2.0 * temp2)
         end
      else
         critcompbetabinomial = i + 1.0
         Exit Function
      end
   end
end

function crit_BetaBinomial(sample_size::Float64, beta_shape1::Float64, beta_shape2::Float64, crit_prob::Float64)::Float64
   sample_size = AlterForIntegralChecks_Others(sample_size)
   if (beta_shape1 <= 0.0 || beta_shape2 <= 0.0 || sample_size < 0.0)
      crit_BetaBinomial = [#VALUE!]
   elseif (crit_prob < 0.0 || crit_prob > 1.0)
      crit_BetaBinomial = [#VALUE!]
   elseif (crit_prob = 0.0)
      crit_BetaBinomial = [#VALUE!]
   elseif (sample_size = 0.0 || crit_prob = 1.0)
      crit_BetaBinomial = sample_size
   else
      Dim i::Float64, pr::Float64
      i = critbetabinomial(beta_shape1, beta_shape2, sample_size, crit_prob)
      crit_BetaBinomial = i
      pr = cdf_BetaBinomial(i, sample_size, beta_shape1, beta_shape2)
      if (pr = crit_prob)
      elseif (pr > crit_prob)
         i = i - 1.0
         pr = cdf_BetaBinomial(i, sample_size, beta_shape1, beta_shape2)
         if (pr >= crit_prob)
            crit_BetaBinomial = i
         end
      else
         crit_BetaBinomial = i + 1.0
      end
   end
   crit_BetaBinomial = GetRidOfMinusZeroes(crit_BetaBinomial)
end

function comp_crit_BetaBinomial(sample_size::Float64, beta_shape1::Float64, beta_shape2::Float64, crit_prob::Float64)::Float64
   sample_size = AlterForIntegralChecks_Others(sample_size)
   if (beta_shape1 <= 0.0 || beta_shape2 <= 0.0 || sample_size < 0.0)
      comp_crit_BetaBinomial = [#VALUE!]
   elseif (crit_prob < 0.0 || crit_prob > 1.0)
      comp_crit_BetaBinomial = [#VALUE!]
   elseif (crit_prob = 1.0)
      comp_crit_BetaBinomial = 0.0
   elseif (sample_size = 0.0 || crit_prob = 0.0)
      comp_crit_BetaBinomial = sample_size
   else
      Dim i::Float64, pr::Float64
      i = critcompbetabinomial(beta_shape1, beta_shape2, sample_size, crit_prob)
      comp_crit_BetaBinomial = i
      pr = comp_cdf_BetaBinomial(i, sample_size, beta_shape1, beta_shape2)
      if (pr = crit_prob)
      elseif (pr < crit_prob)
         i = i - 1.0
         pr = comp_cdf_BetaBinomial(i, sample_size, beta_shape1, beta_shape2)
         if (pr <= crit_prob)
            comp_crit_BetaBinomial = i
         end
      else
         comp_crit_BetaBinomial = i + 1.0
      end
   end
   comp_crit_BetaBinomial = GetRidOfMinusZeroes(comp_crit_BetaBinomial)
end

function pdf_normal_os(x::Float64, Optional n::Float64 = 1, Optional r::Float64 = -1)::Float64
 # pdf for rth smallest (r>0) or -rth largest (r<0) of the n order statistics from a sample of iid N(0,1) variables
 # based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    if n < 1 || abs(r) > n || r = 0 pdf_normal_os = [#VALUE!]: Exit Function
    Dim n1::Float64: n1 = n + 1
    if r > 0 r = r - n1
    if x <= 0
        pdf_normal_os = pdf_beta(cnormal(x), n1 + r, -r) * pdf_normal(x)
    else
        pdf_normal_os = pdf_beta(cnormal(-x), -r, n1 + r) * pdf_normal(-x)
    end
end
 
function cdf_normal_os(x::Float64, Optional n::Float64 = 1, Optional r::Float64 = -1)::Float64
 # cdf for rth smallest (r>0) or -rth largest (r<0) of the n order statistics from a sample of iid N(0,1) variables
 # based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    if n < 1 || abs(r) > n || r = 0 cdf_normal_os = [#VALUE!]: Exit Function
    Dim n1::Float64: n1 = n + 1
    if r > 0 r = r - n1
    if x <= 0
        cdf_normal_os = cdf_beta(cnormal(x), n1 + r, -r)
    else
        cdf_normal_os = comp_cdf_beta(cnormal(-x), -r, n1 + r)
    end
end
 
function comp_cdf_normal_os(x::Float64, Optional n::Float64 = 1, Optional r::Float64 = -1)::Float64
 # 1-cdf for rth smallest (r>0) or -rth largest (r<0) of the n order statistics from a sample of iid N(0,1) variables
 # based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    if n < 1 || abs(r) > n || r = 0 comp_cdf_normal_os = [#VALUE!]: Exit Function
    Dim n1::Float64: n1 = n + 1
    if r > 0 r = r - n1
    if x <= 0
        comp_cdf_normal_os = comp_cdf_beta(cnormal(x), n1 + r, -r)
    else
        comp_cdf_normal_os = cdf_beta(cnormal(-x), -r, n1 + r)
    end
end
 
function inv_normal_os(p::Float64, Optional n::Float64 = 1, Optional r::Float64 = -1)::Float64
 # inverse of cdf_normal_os
 # based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
 # accuracy for median of extreme order statistic is limited by accuracy of IEEE double precision representation of n >> 10^15, not by this routine
    Dim oneMinusxp::Float64
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    if n < 1 || abs(r) > n || r = 0 inv_normal_os = [#VALUE!]: Exit Function
    Dim xp::Float64, n1::Float64: n1 = n + 1
    if r > 0 r = r - n1
    xp = invbeta(n1 + r, -r, p, oneMinusxp)
    if abs(xp - 0.5) < 0.00000000000001 && xp <> 0.5 if cdf_beta(0.5, n1 + r, -r) = p inv_normal_os = 0: Exit Function
    if xp <= 0.5
        inv_normal_os = inv_normal(xp)
    else
        inv_normal_os = -inv_normal(oneMinusxp)
    end
end
 
function comp_inv_normal_os(p::Float64, Optional n::Float64 = 1, Optional r::Float64 = -1)::Float64
 # inverse of comp_cdf_normal_os
 # based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
    Dim oneMinusxp::Float64
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    if n < 1 || abs(r) > n || r = 0 comp_inv_normal_os = [#VALUE!]: Exit Function
    Dim xp::Float64, n1::Float64: n1 = n + 1
    if r > 0 r = r - n1
    xp = invcompbeta(n1 + r, -r, p, oneMinusxp)
    if abs(xp - 0.5) < 0.00000000000001 && xp <> 0.5 if comp_cdf_beta(0.5, n1 + r, -r) = p comp_inv_normal_os = 0: Exit Function
    if xp <= 0.5
        comp_inv_normal_os = inv_normal(xp)
    else
        comp_inv_normal_os = -inv_normal(oneMinusxp)
    end
end

function pdf_gamma_os(x::Float64, shape_param::Float64, Optional n::Float64 = 1, Optional r::Float64 = -1, Optional scale_param::Float64 = 1, Optional nc_param::Float64 = 0)::Float64
 # pdf for rth smallest (r>0) or -rth largest (r<0) of the n order statistics from a sample of iid gamma(a,1) variables
 # based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    if n < 1 || abs(r) > n || r = 0 pdf_gamma_os = [#VALUE!]: Exit Function
    Dim p::Float64, n1::Float64: n1 = n + 1
    if r > 0 r = r - n1
    p = cdf_gamma_nc(x / scale_param, shape_param, nc_param)
    if p <= 0.5        # avoid truncation error by working with p <= 0.5
        pdf_gamma_os = pdf_beta(p, n1 + r, -r) * pdf_gamma_nc(x / scale_param, shape_param, nc_param) / scale_param
    else
        pdf_gamma_os = pdf_beta(comp_cdf_gamma_nc(x / scale_param, shape_param, nc_param), -r, n1 + r) * pdf_gamma_nc(x / scale_param, shape_param, nc_param) / scale_param
    end
end

function cdf_gamma_os(x::Float64, shape_param::Float64, Optional n::Float64 = 1, Optional r::Float64 = -1, Optional scale_param::Float64 = 1, Optional nc_param::Float64 = 0)::Float64
 # cdf for rth smallest (r>0) or -rth largest (r<0) of the n order statistics from a sample of iid gamma(a,1) variables
 # based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    if n < 1 || abs(r) > n || r = 0 cdf_gamma_os = [#VALUE!]: Exit Function
    Dim p::Float64, n1::Float64: n1 = n + 1
    if r > 0 r = r - n1
    p = cdf_gamma_nc(x / scale_param, shape_param, nc_param)
    if p <= 0.5        # avoid truncation error by working with p <= 0.5
        cdf_gamma_os = cdf_beta(p, n1 + r, -r)
    else
        cdf_gamma_os = comp_cdf_beta(comp_cdf_gamma_nc(x / scale_param, shape_param, nc_param), -r, n1 + r)
    end
end

function comp_cdf_gamma_os(x::Float64, shape_param::Float64, Optional n::Float64 = 1, Optional r::Float64 = -1, Optional scale_param::Float64 = 1, Optional nc_param::Float64 = 0)::Float64
 # 1-cdf for rth smallest (r>0) or -rth largest (r<0) of the n order statistics from a sample of iid gamma(a,1) variables
 # based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    if n < 1 || abs(r) > n || r = 0 comp_cdf_gamma_os = [#VALUE!]: Exit Function
    Dim p::Float64, n1::Float64: n1 = n + 1
    if r > 0 r = r - n1
    p = cdf_gamma_nc(x / scale_param, shape_param, nc_param)
    if p <= 0.5        # avoid truncation error by working with p <= 0.5
        comp_cdf_gamma_os = comp_cdf_beta(p, n1 + r, -r)
    else
        comp_cdf_gamma_os = cdf_beta(comp_cdf_gamma_nc(x / scale_param, shape_param, nc_param), -r, n1 + r)
    end
end

function inv_gamma_os(p::Float64, shape_param::Float64, Optional n::Float64 = 1, Optional r::Float64 = -1, Optional scale_param::Float64 = 1, Optional nc_param::Float64 = 0)::Float64
 # inverse of cdf_gamma_os
 # based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
 # accuracy for median of extreme order statistic is limited by accuracy of IEEE double precision representation of n >> 10^15, not by this routine
    Dim oneMinusxp::Float64
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    if n < 1 || abs(r) > n || r = 0 inv_gamma_os = [#VALUE!]: Exit Function
    Dim xp::Float64, n1::Float64: n1 = n + 1
    if r > 0 r = r - n1
    xp = invbeta(n1 + r, -r, p, oneMinusxp)
    if xp <= 0.5       # avoid truncation error by working with xp <= 0.5
        inv_gamma_os = inv_gamma_nc(xp, shape_param, nc_param) * scale_param
    else
        inv_gamma_os = comp_inv_gamma_nc(oneMinusxp, shape_param, nc_param) * scale_param
    end
end

function comp_inv_gamma_os(p::Float64, shape_param::Float64, Optional n::Float64 = 1, Optional r::Float64 = -1, Optional scale_param::Float64 = 1, Optional nc_param::Float64 = 0)::Float64
 # inverse of comp_cdf_gamma_os
 # based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
    Dim oneMinusxp::Float64
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    if n < 1 || abs(r) > n || r = 0 comp_inv_gamma_os = [#VALUE!]: Exit Function
    Dim xp::Float64, n1::Float64: n1 = n + 1
    if r > 0 r = r - n1
    xp = invcompbeta(n1 + r, -r, p, oneMinusxp)
    if xp <= 0.5       # avoid truncation error by working with xp <= 0.5
        comp_inv_gamma_os = inv_gamma_nc(xp, shape_param, nc_param) * scale_param
    else
        comp_inv_gamma_os = comp_inv_gamma_nc(oneMinusxp, shape_param, nc_param) * scale_param
    end
end

function pdf_chi2_os(x::Float64, df::Float64, Optional n::Float64 = 1, Optional r::Float64 = -1, Optional nc_param::Float64 = 0)::Float64
 # pdf for rth smallest (r>0) or -rth largest (r<0) of the n order statistics from a sample of iid chi2(df) variables
 # based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    if n < 1 || abs(r) > n || r = 0 pdf_chi2_os = [#VALUE!]: Exit Function
   df = AlterForIntegralChecks_df(df)
    Dim p::Float64, n1::Float64: n1 = n + 1
    if r > 0 r = r - n1
    p = cdf_Chi2_nc(x, df, nc_param)
    if p <= 0.5        # avoid truncation error by working with p <= 0.5
        pdf_chi2_os = pdf_beta(p, n1 + r, -r) * pdf_Chi2_nc(x, df, nc_param)
    else
        pdf_chi2_os = pdf_beta(comp_cdf_Chi2_nc(x, df, nc_param), -r, n1 + r) * pdf_Chi2_nc(x, df, nc_param)
    end
end

function cdf_chi2_os(x::Float64, df::Float64, Optional n::Float64 = 1, Optional r::Float64 = -1, Optional nc_param::Float64 = 0)::Float64
 # cdf for rth smallest (r>0) or -rth largest (r<0) of the n order statistics from a sample of iid chi2(df) variables
 # based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    if n < 1 || abs(r) > n || r = 0 cdf_chi2_os = [#VALUE!]: Exit Function
   df = AlterForIntegralChecks_df(df)
    Dim p::Float64, n1::Float64: n1 = n + 1
    if r > 0 r = r - n1
    p = cdf_Chi2_nc(x, df, nc_param)
    if p <= 0.5        # avoid truncation error by working with p <= 0.5
        cdf_chi2_os = cdf_beta(p, n1 + r, -r)
    else
        cdf_chi2_os = comp_cdf_beta(comp_cdf_Chi2_nc(x, df, nc_param), -r, n1 + r)
    end
end

function comp_cdf_chi2_os(x::Float64, df::Float64, Optional n::Float64 = 1, Optional r::Float64 = -1, Optional nc_param::Float64 = 0)::Float64
 # 1-cdf for rth smallest (r>0) or -rth largest (r<0) of the n order statistics from a sample of iid chi2(df) variables
 # based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    if n < 1 || abs(r) > n || r = 0 comp_cdf_chi2_os = [#VALUE!]: Exit Function
   df = AlterForIntegralChecks_df(df)
    Dim p::Float64, n1::Float64: n1 = n + 1
    if r > 0 r = r - n1
    p = cdf_Chi2_nc(x, df, nc_param)
    if p <= 0.5        # avoid truncation error by working with p <= 0.5
        comp_cdf_chi2_os = comp_cdf_beta(p, n1 + r, -r)
    else
        comp_cdf_chi2_os = cdf_beta(comp_cdf_Chi2_nc(x, df, nc_param), -r, n1 + r)
    end
end

function inv_chi2_os(p::Float64, df::Float64, Optional n::Float64 = 1, Optional r::Float64 = -1, Optional nc_param::Float64 = 0)::Float64
 # inverse of cdf_chi2_os
 # based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
 # accuracy for median of extreme order statistic is limited by accuracy of IEEE double precision representation of n >> 10^15, not by this routine
    Dim oneMinusxp::Float64
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    if n < 1 || abs(r) > n || r = 0 inv_chi2_os = [#VALUE!]: Exit Function
   df = AlterForIntegralChecks_df(df)
    Dim xp::Float64, n1::Float64: n1 = n + 1
    if r > 0 r = r - n1
    xp = invbeta(n1 + r, -r, p, oneMinusxp)
    if xp <= 0.5       # avoid truncation error by working with xp <= 0.5
        inv_chi2_os = inv_Chi2_nc(xp, df, nc_param)
    else
        inv_chi2_os = comp_inv_Chi2_nc(oneMinusxp, df, nc_param)
    end
end

function comp_inv_chi2_os(p::Float64, df::Float64, Optional n::Float64 = 1, Optional r::Float64 = -1, Optional nc_param::Float64 = 0)::Float64
 # inverse of comp_cdf_chi2_os
 # based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
    Dim oneMinusxp::Float64
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    if n < 1 || abs(r) > n || r = 0 comp_inv_chi2_os = [#VALUE!]: Exit Function
   df = AlterForIntegralChecks_df(df)
    Dim xp::Float64, n1::Float64: n1 = n + 1
    if r > 0 r = r - n1
    xp = invcompbeta(n1 + r, -r, p, oneMinusxp)
    if xp <= 0.5       # avoid truncation error by working with xp <= 0.5
        comp_inv_chi2_os = inv_Chi2_nc(xp, df, nc_param)
    else
        comp_inv_chi2_os = comp_inv_Chi2_nc(oneMinusxp, df, nc_param)
    end
end

function pdf_F_os(x::Float64, df1::Float64, df2::Float64, Optional n::Float64 = 1, Optional r::Float64 = -1, Optional nc_param::Float64 = 0)::Float64
 # pdf for rth smallest (r>0) or -rth largest (r<0) of the n order statistics from a sample of iid gamma(a,1) variables
 # based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    if n < 1 || abs(r) > n || r = 0 pdf_F_os = [#VALUE!]: Exit Function
   df1 = AlterForIntegralChecks_df(df1)
   df2 = AlterForIntegralChecks_df(df2)
    Dim p::Float64, n1::Float64: n1 = n + 1
    if r > 0 r = r - n1
    p = cdf_fdist_nc(x, df1, df2, nc_param)
    if p <= 0.5        # avoid truncation error by working with p <= 0.5
        pdf_F_os = pdf_beta(p, n1 + r, -r) * pdf_fdist_nc(x, df1, df2, nc_param)
    else
        pdf_F_os = pdf_beta(comp_cdf_fdist_nc(x, df1, df2, nc_param), -r, n1 + r) * pdf_fdist_nc(x, df1, df2, nc_param)
    end
end

function cdf_F_os(x::Float64, df1::Float64, df2::Float64, Optional n::Float64 = 1, Optional r::Float64 = -1, Optional nc_param::Float64 = 0)::Float64
 # cdf for rth smallest (r>0) or -rth largest (r<0) of the n order statistics from a sample of iid gamma(a,1) variables
 # based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    if n < 1 || abs(r) > n || r = 0 cdf_F_os = [#VALUE!]: Exit Function
   df1 = AlterForIntegralChecks_df(df1)
   df2 = AlterForIntegralChecks_df(df2)
    Dim p::Float64, n1::Float64: n1 = n + 1
    if r > 0 r = r - n1
    p = cdf_fdist_nc(x, df1, df2, nc_param)
    if p <= 0.5        # avoid truncation error by working with p <= 0.5
        cdf_F_os = cdf_beta(p, n1 + r, -r)
    else
        cdf_F_os = comp_cdf_beta(comp_cdf_fdist_nc(x, df1, df2, nc_param), -r, n1 + r)
    end
end

function comp_cdf_F_os(x::Float64, df1::Float64, df2::Float64, Optional n::Float64 = 1, Optional r::Float64 = -1, Optional nc_param::Float64 = 0)::Float64
 # 1-cdf for rth smallest (r>0) or -rth largest (r<0) of the n order statistics from a sample of iid gamma(a,1) variables
 # based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    if n < 1 || abs(r) > n || r = 0 comp_cdf_F_os = [#VALUE!]: Exit Function
   df1 = AlterForIntegralChecks_df(df1)
   df2 = AlterForIntegralChecks_df(df2)
    Dim p::Float64, n1::Float64: n1 = n + 1
    if r > 0 r = r - n1
    p = cdf_fdist_nc(x, df1, df2, nc_param)
    if p <= 0.5        # avoid truncation error by working with p <= 0.5
        comp_cdf_F_os = comp_cdf_beta(p, n1 + r, -r)
    else
        comp_cdf_F_os = cdf_beta(comp_cdf_fdist_nc(x, df1, df2, nc_param), -r, n1 + r)
    end
end

function inv_F_os(p::Float64, df1::Float64, df2::Float64, Optional n::Float64 = 1, Optional r::Float64 = -1, Optional nc_param::Float64 = 0)::Float64
 # inverse of cdf_F_os
 # based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
 # accuracy for median of extreme order statistic is limited by accuracy of IEEE double precision representation of n >> 10^15, not by this routine
    Dim oneMinusxp::Float64
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    if n < 1 || abs(r) > n || r = 0 inv_F_os = [#VALUE!]: Exit Function
   df1 = AlterForIntegralChecks_df(df1)
   df2 = AlterForIntegralChecks_df(df2)
    Dim xp::Float64, n1::Float64: n1 = n + 1
    if r > 0 r = r - n1
    xp = invbeta(n1 + r, -r, p, oneMinusxp)
    if xp <= 0.5       # avoid truncation error by working with xp <= 0.5
        inv_F_os = inv_fdist_nc(xp, df1, df2, nc_param)
    else
        inv_F_os = comp_inv_fdist_nc(oneMinusxp, df1, df2, nc_param)
    end
end

function comp_inv_F_os(p::Float64, df1::Float64, df2::Float64, Optional n::Float64 = 1, Optional r::Float64 = -1, Optional nc_param::Float64 = 0)::Float64
 # inverse of comp_cdf_F_os
 # based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
    Dim oneMinusxp::Float64
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    if n < 1 || abs(r) > n || r = 0 comp_inv_F_os = [#VALUE!]: Exit Function
   df1 = AlterForIntegralChecks_df(df1)
   df2 = AlterForIntegralChecks_df(df2)
    Dim xp::Float64, n1::Float64: n1 = n + 1
    if r > 0 r = r - n1
    xp = invcompbeta(n1 + r, -r, p, oneMinusxp)
    if xp <= 0.5       # avoid truncation error by working with xp <= 0.5
        comp_inv_F_os = inv_fdist_nc(xp, df1, df2, nc_param)
    else
        comp_inv_F_os = comp_inv_fdist_nc(oneMinusxp, df1, df2, nc_param)
    end
end

function pdf_beta_os(x::Float64, shape_param1::Float64, shape_param2::Float64, Optional n::Float64 = 1, Optional r::Float64 = -1, Optional nc_param::Float64 = 0)::Float64
 # pdf for rth smallest (r>0) or -rth largest (r<0) of the n order statistics from a sample of iid gamma(a,1) variables
 # based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    if n < 1 || abs(r) > n || r = 0 pdf_beta_os = [#VALUE!]: Exit Function
    Dim p::Float64, n1::Float64: n1 = n + 1
    if r > 0 r = r - n1
    p = cdf_beta_nc(x, shape_param1, shape_param2, nc_param)
    if p <= 0.5        # avoid truncation error by working with p <= 0.5
        pdf_beta_os = pdf_beta(p, n1 + r, -r) * pdf_beta_nc(x, shape_param1, shape_param2, nc_param)
    else
        pdf_beta_os = pdf_beta(comp_cdf_beta_nc(x, shape_param1, shape_param2, nc_param), -r, n1 + r) * pdf_beta_nc(x, shape_param1, shape_param2, nc_param)
    end
end

function cdf_beta_os(x::Float64, shape_param1::Float64, shape_param2::Float64, Optional n::Float64 = 1, Optional r::Float64 = -1, Optional nc_param::Float64 = 0)::Float64
 # cdf for rth smallest (r>0) or -rth largest (r<0) of the n order statistics from a sample of iid gamma(a,1) variables
 # based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    if n < 1 || abs(r) > n || r = 0 cdf_beta_os = [#VALUE!]: Exit Function
    Dim p::Float64, n1::Float64: n1 = n + 1
    if r > 0 r = r - n1
    p = cdf_beta_nc(x, shape_param1, shape_param2, nc_param)
    if p <= 0.5        # avoid truncation error by working with p <= 0.5
        cdf_beta_os = cdf_beta(p, n1 + r, -r)
    else
        cdf_beta_os = comp_cdf_beta(comp_cdf_beta_nc(x, shape_param1, shape_param2, nc_param), -r, n1 + r)
    end
end

function comp_cdf_beta_os(x::Float64, shape_param1::Float64, shape_param2::Float64, Optional n::Float64 = 1, Optional r::Float64 = -1, Optional nc_param::Float64 = 0)::Float64
 # 1-cdf for rth smallest (r>0) or -rth largest (r<0) of the n order statistics from a sample of iid gamma(a,1) variables
 # based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    if n < 1 || abs(r) > n || r = 0 comp_cdf_beta_os = [#VALUE!]: Exit Function
    Dim p::Float64, n1::Float64: n1 = n + 1
    if r > 0 r = r - n1
    p = cdf_beta_nc(x, shape_param1, shape_param2, nc_param)
    if p <= 0.5        # avoid truncation error by working with p <= 0.5
        comp_cdf_beta_os = comp_cdf_beta(p, n1 + r, -r)
    else
        comp_cdf_beta_os = cdf_beta(comp_cdf_beta_nc(x, shape_param1, shape_param2, nc_param), -r, n1 + r)
    end
end

function inv_beta_os(p::Float64, shape_param1::Float64, shape_param2::Float64, Optional n::Float64 = 1, Optional r::Float64 = -1, Optional nc_param::Float64 = 0)::Float64
 # inverse of cdf_beta_os
 # based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
 # accuracy for median of extreme order statistic is limited by accuracy of IEEE double precision representation of n >> 10^15, not by this routine
    Dim oneMinusxp::Float64
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    if n < 1 || abs(r) > n || r = 0 inv_beta_os = [#VALUE!]: Exit Function
    Dim xp::Float64, n1::Float64: n1 = n + 1
    if r > 0 r = r - n1
    xp = invbeta(n1 + r, -r, p, oneMinusxp)
    if xp <= 0.5       # avoid truncation error by working with xp <= 0.5
        inv_beta_os = inv_beta_nc(xp, shape_param1, shape_param2, nc_param)
    else
        inv_beta_os = comp_inv_beta_nc(oneMinusxp, shape_param1, shape_param2, nc_param)
    end
end

function comp_inv_beta_os(p::Float64, shape_param1::Float64, shape_param2::Float64, Optional n::Float64 = 1, Optional r::Float64 = -1, Optional nc_param::Float64 = 0)::Float64
 # inverse of comp_cdf_beta_os
 # based on formula 2.1.5 "Order Statistics" by H.A. David (any edition)
    Dim oneMinusxp::Float64
    n = AlterForIntegralChecks_Others(n): r = AlterForIntegralChecks_Others(r)
    if n < 1 || abs(r) > n || r = 0 comp_inv_beta_os = [#VALUE!]: Exit Function
    Dim xp::Float64, n1::Float64: n1 = n + 1
    if r > 0 r = r - n1
    xp = invcompbeta(n1 + r, -r, p, oneMinusxp)
    if xp <= 0.5       # avoid truncation error by working with xp <= 0.5
        comp_inv_beta_os = inv_beta_nc(xp, shape_param1, shape_param2, nc_param)
    else
        comp_inv_beta_os = comp_inv_beta_nc(oneMinusxp, shape_param1, shape_param2, nc_param)
    end
end

function fet_pearson22(a::Float64, b::Float64, c::Float64, d::Float64)::Float64
#The following is some VBA code for the two-sided 2x2 FET based on Pearson#s
#Chi-Square statistic (i.e. includes all tables which give a value of the
#Chi-Square statistic which is greater than or equal to that of the table observed)
Dim det::Float64, temp::Float64, sample_size::Float64, pop::Float64
det = a * d - b * c
if det > 0
    temp = a
    a = b
    b = temp
    temp = c
    c = d
    d = temp
    det = -det
end
sample_size = a + b
temp = a + c
pop = sample_size + c + d
det = (2.0 * det + 1) / pop
if det < -1.0
   fet_pearson22 = cdf_hypergeometric(a, sample_size, temp, pop) + comp_cdf_hypergeometric(a - det, sample_size, temp, pop)
else
   fet_pearson22 = 1.0
end
end

function chi_square_test(r As Range)::Float64
Dim cs::Float64, rs::Float64
Dim rc As Long, cc As Long, i As Long, j As Long, k As Long
rc = r.Rows.count
cc = r.Columns.count
if rc < 2 || cc < 2
   chi_square_test = [#VALUE!]
   Exit Function
end
ReDim os(1 To rc, 1 To cc)::Float64, Es(0 To rc, 0 To cc)::Float64
For i = 1 To rc
    For j = 1 To cc
        os(i, j) = r.Item(i, j)
    Next j
Next i
#Calculate row totals and check that all values are non-negative integers
cs = 0.0
For i = 1 To rc
   rs = 0.0
   For j = 1 To cc
      if os(i, j) < 0 || Int(os(i, j)) <> os(i, j)
         chi_square_test = [#VALUE!]
         Exit Function
      end
      rs = rs + os(i, j)
   Next j
   Es(i, 0) = rs
   cs = cs + rs
Next i
Es(0, 0) = cs
#Calculate column totals
For i = 1 To cc
   rs = 0.0
   For j = 1 To rc
      rs = rs + os(j, i)
   Next j
   Es(0, i) = rs
Next i
#Calculate chi_square value
rs = 0.0
For i = 1 To rc
   For j = 1 To cc
      Es(i, j) = Es(i, 0) * Es(0, j) / cs
      rs = rs + (os(i, j) - Es(i, j)) ^ 2 / Es(i, j)
   Next j
Next i
chi_square_test = comp_cdf_chi_sq(rs, (rc - 1) * (cc - 1))

end

function nidf_fdist(x::Float64, df1::Float64, df2::Float64)::Float64
   if (df1 <= 0.0 || df2 <= 0.0)
      nidf_fdist = [#VALUE!]
   elseif (x <= 0.0)
      nidf_fdist = 0.0
   else
      Dim p::Float64, q::Float64
      p = df1 * x
      q = df2 + p
      p = p / q
      q = df2 / q
      df2 = df2 / 2.0
      df1 = df1 / 2.0
      if (p < 0.5)
          nidf_fdist = beta(p, df1, df2)
      else
          nidf_fdist = compbeta(q, df2, df1)
      end
   end
end

function comp_nidf_fdist(x::Float64, df1::Float64, df2::Float64)::Float64
   if (df1 <= 0.0 || df2 <= 0.0)
      comp_nidf_fdist = [#VALUE!]
   elseif (x <= 0.0)
      comp_nidf_fdist = 1.0
   else
      Dim p::Float64, q::Float64
      p = df1 * x
      q = df2 + p
      p = p / q
      q = df2 / q
      df2 = df2 / 2.0
      df1 = df1 / 2.0
      if (p < 0.5)
          comp_nidf_fdist = compbeta(p, df1, df2)
      else
          comp_nidf_fdist = beta(q, df2, df1)
      end
   end
end

function CBNB(i::Float64, r::Float64, beta_shape1::Float64, beta_shape2::Float64)::Float64
   Dim j::Float64
   j = Int(i)
   CBNB = 0.0
   while j > -1
      CBNB = CBNB + pmf_BetaNegativeBinomial(j, r, beta_shape1, beta_shape2)
      j = j - 1
   end
end

function CBNB1(i::Float64, r::Float64, beta_shape1::Float64, beta_shape2::Float64)::Float64
   Dim j::Float64, ss::Float64, bs2::Float64, temp::Float64, swap::Float64
   On Error GoTo errorhandler
   j = Int(i)
   ss = min(r, beta_shape2)
   bs2 = max(r, beta_shape2)
   CBNB1 = 0.0
   temp = 0.0
   if beta_shape1 >= 1.0
      beta_shape1 = beta_shape1 - 1.0
      temp = hypergeometricTerm(j, bs2, ss, beta_shape1)
      CBNB1 = temp
      while (beta_shape1 >= 1.0) && (temp > 1E-16 * CBNB1)
         bs2 = bs2 + 1.0
         ss = ss + 1.0
         temp = temp * (j * beta_shape1) / (bs2 * ss)
         j = j - 1.0
         beta_shape1 = beta_shape1 - 1.0
         CBNB1 = CBNB1 + temp
      end
      if beta_shape1 >= 1.0 Exit Function
      j = j - 1.0
      bs2 = bs2 + 1.0
      ss = ss + 1.0
   end
   if temp = 0.0
      temp = PBB(j, ss, bs2, beta_shape1)
   else
      temp = temp * ((j + 1) * beta_shape1) / (ss * (j + bs2))
   end
   CBNB1 = CBNB1 + temp
   while j > 0
      temp = (j * (ss + beta_shape1)) * temp
      ss = ss + 1
      j = j - 1
      if ss > bs2
         swap = ss
         ss = bs2
         bs2 = swap
      end
      temp = temp / ((j + bs2) * ss)
      CBNB1 = CBNB1 + temp
      if temp < 1E-16 * CBNB1 Exit Do
   end
#Debug.Print j, ss, bs2, beta_shape1
   Exit Function
errorhandler: Debug.Print j, ss, bs2, beta_shape1
end

function ccBNB(i::Float64, r::Float64, beta_shape1::Float64, beta_shape2::Float64)::Float64
   Dim ha1::Float64, hprob::Float64, hswap::Bool
   Dim mrb2::Float64, other::Float64, temp::Float64, ctemp::Float64, mnib1::Float64, mxib1::Float64, swap::Float64, max_iterations::Float64
   i = Int(i)
   mrb2 = max(r, beta_shape2)
   other = min(r, beta_shape2)
   if (other <= 0.0 || beta_shape1 <= 0.0)
      ccBNB = [#VALUE!]
   elseif i < 0.0
      ccBNB = 1.0
   elseif (mrb2 >= 100.0 || (mrb2 >= 5.0 && other * mrb2 > (i + 0.5) * (beta_shape1 + 0.5)))
      ccBNB = CBNB0(mrb2 - 1.0, i + 1.0, other, beta_shape1, 0.0)
   else
      mxib1 = beta_shape1
      mnib1 = i + 1.0
      temp = pmf_BetaNegativeBinomial(mnib1, r, mxib1, beta_shape2)
      ctemp = temp
      max_iterations = max(60.0, mrb2)
      while (mnib1 + mxib1 < max_iterations) && temp > 1E-16 * ctemp
         temp = temp * (mnib1 + r) * (mnib1 + beta_shape2) / (mnib1 + mxib1 + r + beta_shape2)
         mnib1 = mnib1 + 1.0
         if mxib1 < mnib1
            swap = mxib1
            mxib1 = mnib1
            mnib1 = swap
         end
         temp = temp / mnib1
         ctemp = ctemp + temp
      end
      temp = hypergeometricTerm(r, mnib1, mxib1, beta_shape2) * (r * beta_shape2 * (mnib1 + r + mxib1 + beta_shape2)) / ((mnib1 + 1.0) * (mxib1 + r) * (mxib1 + beta_shape2))
      ctemp = ctemp + temp
      mnib1 = mnib1 + 1.0
      r = r - 1.0
      beta_shape2 = beta_shape2 - 1.0
      Do
         mnib1 = mnib1 + 1.0
         mxib1 = mxib1 + 1.0
         temp = temp * (r * beta_shape2) / (mnib1 * mxib1)
         r = r - 1.0
         beta_shape2 = beta_shape2 - 1.0
         ctemp = ctemp + temp
      end Until abs(temp) <= 1E-16 * ctemp
      ccBNB = ctemp
   end
   ccBNB = GetRidOfMinusZeroes(ccBNB)
end

Function ccBNB5(ilim::Float64, rr::Float64, a::Float64, bb::Float64)::Float64
   Dim temp::Float64, i::Float64, r::Float64, b::Float64
   if rr > bb
      r = rr
      b = bb
   else
      r = bb
      b = rr
   end
   ccBNB5 = (a + 0.5) * log0(b * r / ((a + 1.0) * (b + a + r + 1.0))) - r * log0(b / (r + a + 1.0)) - b * log0(r / (a + b + 1.0))
   if r <= 0.001
      temp = a + (b + r) * 0.5
      ccBNB5 = ccBNB5  - b * r * (logfbit2(temp) + (b ^ 2 + r ^ 2) * logfbit4(temp) / 24.0)
   else
      ccBNB5 = ccBNB5  + (lfbaccdif1(b, r + a) - lfbaccdif1(b, a))
   end
   temp = 0.0
   if ilim > 0.0
      i = ilim
      while i > 1.0
         i = i - 1.0
         temp = (1.0 + temp) * (i + r) * (i + b) / ((i + r + a + b) * (i + 1.0))
      end
      temp = (1.0 + temp) * exp(ccBNB5) * a
   end
   ccBNB5 = (r * b * (1.0 - temp) - expm1(ccBNB5) * a * (r + a + b)) / ((r + a) * (a + b))
end

Function fet_22(c As Long, colsum()::Float64, rowsum::Float64, pmf_Obs::Float64, ByRef inumstart::Float64, ByRef jnumstart::Float64)::Float64
#The following is some VBA code for the two-sided 2x2 FET based on Pearson#s
#Chi-Square statistic (i.e. includes all tables which give a value of the
#Chi-Square statistic which is greater than or equal to that of the table observed)
Dim inum_min::Float64, jnum_max::Float64, pmf_table::Float64, inum::Float64, jnum::Float64, mode::Float64, pmrc::Float64, pop::Float64, d::Float64, prob_d::Float64, pmfh::Float64, pmfh_save::Float64, knum::Float64, prob::Float64
Dim i As Long, j As Long
Dim all_d_zero::Bool
ReDim ml(1 To c)::Float64

#c = High(colsum) #But can#t pass partial arrays in calls
if pmf_Obs >= 1.0  #All tables have pmf <= 1
   fet_22 = 1.0
   Exit Function
end
pop = 0.0
For i = 1 To c
   pop = pop + colsum(i)
Next i

if c = 2
   pmrc = pop - rowsum - colsum(2)
   mode = Int((rowsum + 1.0) * (colsum(2) + 1.0) / (pop + 2.0))
   inum_min = max(0.0, -pmrc)
   inum = min(max(inum_min, inumstart), mode)
   jnum_max = min(rowsum, colsum(2))
   jnum = max(min(jnum_max, jnumstart), mode)
   pmf_table = pmf_hypergeometric(inum, rowsum, colsum(2), pop)
   while pmf_table = 0.0
      inum = Int((inum + max(2, mode)) * 0.5)
      pmf_table = pmf_hypergeometric(inum, rowsum, colsum(2), pop)
   end
   while (pmf_table > pmf_Obs) && (inum > inum_min)
      pmf_table = pmf_table * (inum * (pmrc + inum))
      inum = inum - 1.0
      pmf_table = pmf_table / ((rowsum - inum) * (colsum(2) - inum))
   end
   while (pmf_table <= pmf_Obs) && (inum <= mode)
      pmf_table = pmf_table * ((rowsum - inum) * (colsum(2) - inum))
      inum = inum + 1.0
      pmf_table = pmf_table / (inum * (pmrc + inum))
   end
   pmf_table = pmf_hypergeometric(jnum, rowsum, colsum(2), pop)
   while pmf_table = 0.0
      jnum = Int((jnum + mode) * 0.5)
      pmf_table = pmf_hypergeometric(jnum, rowsum, colsum(2), pop)
   end
   while (pmf_table > pmf_Obs) && (jnum < jnum_max)
      pmf_table = pmf_table * ((rowsum - jnum) * (colsum(2) - jnum))
      jnum = jnum + 1.0
      pmf_table = pmf_table / (jnum * (pmrc + jnum))
   end
   while (pmf_table < pmf_Obs) && (jnum >= mode)
      pmf_table = pmf_table * (jnum * (pmrc + jnum))
      jnum = jnum - 1.0
      pmf_table = pmf_table / ((rowsum - jnum) * (colsum(2) - jnum))
   end
   inumstart = inum
   jnumstart = jnum
   if inum > jnum
      fet_22 = 1.0
   else
      fet_22 = cdf_hypergeometric(inum - 1, rowsum, colsum(2), pop) + comp_cdf_hypergeometric(jnum, rowsum, colsum(2), pop)
   end
else
#First guess at mode vector
   ml(1) = rowsum
   For i = 2 To c
      ml(i) = Int(rowsum * colsum(i) / pop + 0.5)
      ml(1) = ml(1) - ml(i)
   Next i
   
   Do #Update guess at mode vector
      all_d_zero = true
      For i = 1 To c - 1
         For j = i + 1 To c
            d = ml(i) - Int((colsum(i) + 1.0) * (ml(i) + ml(j) + 1.0) / (colsum(i) + colsum(j) + 2.0))
            if d <> 0.0
               ml(i) = ml(i) - d
               ml(j) = ml(j) + d
               all_d_zero = false
            end
         Next j
      Next i
   end Until all_d_zero
   knum = ml(c)
   pmfh = pmf_hypergeometric(knum, rowsum, colsum(c), pop)
   if pmfh = 0.0  #Not entirely sure what we want here but not likely that many tables have pmf < 1e-4933 and if there are it will be vary slow!
      fet_22 = "Probability of table is 0"
      Exit Function
   end
   pmfh_save = pmfh
   inum = inumstart
   jnum = jnumstart
   prob_d = fet_22(c - 1, colsum, rowsum - knum, pmf_Obs / pmfh, inum, jnum)
#Debug.Print knum, pmfh, prob_d
   inumstart = inum
   jnumstart = jnum
   if inum > jnum
      fet_22 = 1.0
      Exit Function
   end
   prob = pmfh * prob_d
   Do
      pmfh = pmfh * (knum * (pop - colsum(c) - rowsum + knum))
      knum = knum - 1.0
      pmfh = pmfh / ((colsum(c) - knum) * (rowsum - knum))
      if pmfh = 0.0 Exit Do
      prob_d = fet_22(c - 1, colsum, rowsum - knum, pmf_Obs / pmfh, inum, jnum)
#Debug.Print knum, pmfh, prob_d
      if inum > jnum
         prob = prob + cdf_hypergeometric(knum, colsum(c), rowsum, pop)
         Exit Do
      end
      prob = prob + pmfh * prob_d
   end

   pmfh = pmfh_save
   knum = ml(c)
   inum = inumstart
   jnum = jnumstart
   Do
      pmfh = pmfh * ((colsum(c) - knum) * (rowsum - knum))
      knum = knum + 1.0
      pmfh = pmfh / (knum * (pop - colsum(c) - rowsum + knum))
      if pmfh = 0.0 Exit Do
      prob_d = fet_22(c - 1, colsum, rowsum - knum, pmf_Obs / pmfh, inum, jnum)
#Debug.Print knum, pmfh, prob_d
      if inum > jnum
         prob = prob + comp_cdf_hypergeometric(knum - 1.0, colsum(c), rowsum, pop)
         Exit Do
      end
      prob = prob + pmfh * prob_d
   end
   fet_22 = prob
end

end

Function old_fet_22(a::Float64, b::Float64, c::Float64, d::Float64)::Float64
#The following is some VBA code for the two-sided 2x2 FET based on Pearson#s
#Chi-Square statistic (i.e. includes all tables which give a value of the
#Chi-Square statistic which is greater than or equal to that of the table observed)
Dim det::Float64, temp::Float64, sample_size::Float64, pop::Float64, pmf_Obs::Float64, pmf_table::Float64, jnum::Float64, mode::Float64
det = GeneralabMinuscd(a, d, b, c)
if det > 0
temp = a
a = b
b = temp
temp = c
c = d
d = temp
det = -det
end
sample_size = a + b
temp = a + c
pop = sample_size + c + d
det = Int((2 * det + 1) / pop)
pmf_Obs = pmf_hypergeometric(a, sample_size, temp, pop) * (1.0000000000001)
if pmf_Obs = 0.0
   old_fet_22 = 0.0
   Exit Function
end
mode = Int((sample_size + 1.0) * (temp + 1.0) / (pop + 2.0))
jnum = a - det
pmf_table = pmf_hypergeometric(jnum, sample_size, temp, pop)
while pmf_table = 0.0
   jnum = Int((jnum + mode) * 0.5)
   pmf_table = pmf_hypergeometric(jnum, sample_size, temp, pop)
end
if pmf_table > pmf_Obs
   while pmf_table >= pmf_Obs
      pmf_table = pmf_table * ((sample_size - jnum) * (temp - jnum))
      jnum = jnum + 1.0
      pmf_table = pmf_table / (jnum * (pop - sample_size - temp + jnum))
   end
   jnum = jnum - 1.0
else
   while pmf_table <= pmf_Obs && jnum >= mode
      pmf_table = pmf_table * (jnum * (pop - sample_size - temp + jnum))
      jnum = jnum - 1.0
      pmf_table = pmf_table / ((sample_size - jnum) * (temp - jnum))
   end
end
if a > jnum
   old_fet_22 = 1.0
else
   old_fet_22 = cdf_hypergeometric(a, sample_size, temp, pop) + comp_cdf_hypergeometric(jnum, sample_size, temp, pop)
end
end

Function fet_23(c As Long, ByRef colsum()::Float64, rowsum::Float64, pmf_Obs::Float64, ByRef inumstart::Float64, ByRef jnumstart::Float64)::Float64
Dim d::Float64, cs::Float64, colsum12::Float64, prob::Float64, pmfh::Float64, pmfh_save::Float64, temp::Float64, inum::Float64, jnum::Float64, knum::Float64
Dim cdf::Float64, ccdf::Float64, pmf_table::Float64, mode::Float64, cdf_save::Float64, ccdf_save::Float64, col1mRowSum::Float64, prob_d::Float64, htTemp::Float64
Dim pmf_table_inum::Float64, pmf_table_jnum::Float64, pmf_table_inum_save::Float64, pmf_table_jnum_save::Float64
Dim i As Long, j As Long, k As Long
Dim all_d_zero::Bool
Dim ast As TAddStack

ReDim ml(1 To c)::Float64

if pmf_Obs > 1.0 #All tables have pmf <= 1
   fet_23 = 1.0
   Exit Function
end

cs = 0.0
For i = 1 To c
   cs = cs + colsum(i)
Next i
#First guess at mode vector
ml(1) = rowsum
For i = 2 To c
   ml(i) = Int(rowsum * colsum(i) / cs + 0.5)
   ml(1) = ml(1) - ml(i)
Next i

Do #Update guess at mode vector
   all_d_zero = true
   For i = 1 To c - 1
      For j = i + 1 To c
         d = ml(i) - Int((colsum(i) + 1.0) * (ml(i) + ml(j) + 1.0) / (colsum(i) + colsum(j) + 2.0))
         if d <> 0.0
            ml(i) = ml(i) - d
            ml(j) = ml(j) + d
            all_d_zero = false
         end
      Next j
   Next i
end Until all_d_zero
knum = ml(c)
pmfh = pmf_hypergeometric(knum, rowsum, colsum(c), cs)
pmfh_save = pmfh

if c = 3
   colsum12 = colsum(1) + colsum(2)
   col1mRowSum = colsum(1) - rowsum
   inum = max(max(0.0, -(knum + col1mRowSum)), inumstart)
   jnum = min(min(rowsum - knum, colsum(2)), jnumstart)
   mode = Int((rowsum - knum + 1.0) * (colsum(2) + 1.0) / (colsum12 + 2.0))
   if inum > mode
      inum = mode
   end
   if jnum < mode
      jnum = mode
   end
   pmf_table = pmf_hypergeometric(inum, rowsum - knum, colsum(2), colsum12)
   while pmf_table = 0.0
      inum = Int((inum + max(2, mode)) * 0.5)
      pmf_table = pmf_hypergeometric(inum, rowsum - knum, colsum(2), colsum12)
   end
   if pmf_table * pmfh <= pmf_Obs
      Do
         pmf_table = pmf_table * ((rowsum - knum - inum) * (colsum(2) - inum))
         inum = inum + 1.0
         pmf_table = pmf_table / (inum * (col1mRowSum + knum + inum))
      end Until (pmf_table * pmfh > pmf_Obs) || (inum > mode)
      pmf_table_inum = pmf_table
   else
      Do
         pmf_table_inum = pmf_table
         pmf_table = pmf_table * (inum * (col1mRowSum + knum + inum))
         inum = inum - 1.0
         pmf_table = pmf_table / ((rowsum - knum - inum) * (colsum(2) - inum))
      end Until pmf_table * pmfh <= pmf_Obs
      inum = inum + 1.0
   end
   pmf_table = pmf_hypergeometric(jnum, rowsum - knum, colsum(2), colsum12)
   while pmf_table = 0.0
      jnum = Int((jnum + mode) * 0.5)
      pmf_table = pmf_hypergeometric(jnum, rowsum - knum, colsum(2), colsum12)
   end
   if pmf_table * pmfh > pmf_Obs
      Do
         pmf_table_jnum = pmf_table
         pmf_table = pmf_table * ((rowsum - knum - jnum) * (colsum(2) - jnum))
         jnum = jnum + 1.0
         pmf_table = pmf_table / (jnum * (col1mRowSum + knum + jnum))
      end Until pmf_table * pmfh <= pmf_Obs
      jnum = jnum - 1.0
   else
      Do
         pmf_table = pmf_table * (jnum * (col1mRowSum + knum + jnum))
         jnum = jnum - 1.0
         pmf_table = pmf_table / ((rowsum - knum - jnum) * (colsum(2) - jnum))
      end Until (pmf_table * pmfh > pmf_Obs) || (jnum < mode)
      pmf_table_jnum = pmf_table
   end
   inumstart = inum
   jnumstart = jnum
   if inum > jnum
      fet_23 = 1.0
      Exit Function
   end
   pmf_table_inum_save = pmf_table_inum
   pmf_table_jnum_save = pmf_table_jnum
   cdf = cdf_hypergeometric(inum - 1, rowsum - knum, colsum(2), colsum12)
   ccdf = comp_cdf_hypergeometric(jnum, rowsum - knum, colsum(2), colsum12)
   prob = 0.0
   Call InitAddStack(ast)
   Call AddValueToStack(ast, pmfh * (cdf + ccdf))
#Debug.Print knum, pmfh * (cdf + ccdf)
   cdf_save = cdf
   ccdf_save = ccdf
   k = 0
   while knum >= 0
      k = k + 1
      if k = 100
         knum = knum - 1.0
         pmfh = pmf_hypergeometric(knum, rowsum, colsum(3), cs)
         k = 0
      else
         pmfh = pmfh * (knum * (colsum12 - rowsum + knum))
         knum = knum - 1.0
         pmfh = pmfh / ((rowsum - knum) * (colsum(3) - knum))
         #pmfh = pmf_hypergeometric(knum, rowsum, colsum(3), cs)
      end
      if pmfh <= pmf_Obs Exit Do
      mode = Int((rowsum - knum + 1.0) * (colsum(2) + 1.0) / (colsum12 + 2.0))
#if knum = 4294567294.0
#   Debug.Print "Got here"
#end
      inum = inum + 1.0
      pmf_table = pmf_table_inum * ((rowsum - knum) * (colsum(2) - inum + 1.0)) / (inum * (colsum12 - rowsum + knum + 1.0))      #pmf_hypergeometric(inum, rowsum - knum, colsum(2), colsum12)
      temp = pmf_table_inum * (col1mRowSum + knum + inum) / (colsum12 - rowsum + knum + 1.0)   #PBB(inum-1, colsum(2) - inum+1, rowsum - knum - inum+1, col1mRowSum + knum + inum)
      if pmf_table = 0.0
         inum = inum - 1.0
         pmf_table = temp * ((rowsum - knum)) / ((rowsum - knum - inum))  #pmf_hypergeometric(inum, rowsum - knum, colsum(2), colsum12)
         temp = temp - pmf_table
      end
      while inum > mode
         pmf_table = pmf_table * (inum * (col1mRowSum + knum + inum))
         inum = inum - 1.0
         pmf_table = pmf_table / ((rowsum - knum - inum) * (colsum(2) - inum))
         temp = temp - pmf_table
      end
      if pmf_table * pmfh <= pmf_Obs
         Do
            temp = temp + pmf_table
            pmf_table = pmf_table * ((rowsum - knum - inum) * (colsum(2) - inum))
            inum = inum + 1.0
            pmf_table = pmf_table / (inum * (col1mRowSum + knum + inum))
         end Until (pmf_table * pmfh > pmf_Obs) || (inum > mode)
         pmf_table_inum = pmf_table
      else
         Do
            pmf_table_inum = pmf_table
            pmf_table = pmf_table * (inum * (col1mRowSum + knum + inum))
            inum = inum - 1.0
            pmf_table = pmf_table / ((rowsum - knum - inum) * (colsum(2) - inum))
            if pmf_table = 0.0
               cdf = 0.0
               temp = 0.0
               Exit Do
            end
            if pmf_table * pmfh <= pmf_Obs Exit Do
            temp = temp - pmf_table
         end
         inum = inum + 1.0
      end
      if k = 50 pmf_table_inum = pmf_hypergeometric(inum, rowsum - knum, colsum(2), colsum12)
      cdf = cdf + temp
      pmf_table = pmf_table_jnum * ((rowsum - knum) * (col1mRowSum + knum + jnum + 1.0)) / ((rowsum - knum - jnum) * (colsum12 - rowsum + knum + 1.0)) #pmf_hypergeometric(jnum, rowsum - knum, colsum(2), colsum12)
      temp = pmf_table_jnum * (colsum(2) - jnum) / (colsum12 - rowsum + knum + 1.0)
      if pmf_table = 0.0
         jnum = jnum + 1.0
         pmf_table = temp * (rowsum - knum) / jnum #pmf_hypergeometric(jnum, rowsum - knum, colsum(2), colsum12)
         temp = temp - pmf_table
      end
      while jnum < mode
         pmf_table = pmf_table * ((rowsum - knum - jnum) * (colsum(2) - jnum))
         jnum = jnum + 1.0
         pmf_table = pmf_table / (jnum * (col1mRowSum + knum + jnum))
         temp = temp - pmf_table
      end
      if pmf_table * pmfh <= pmf_Obs
         Do
            temp = temp + pmf_table
            pmf_table = pmf_table * (jnum * (col1mRowSum + knum + jnum))
            jnum = jnum - 1.0
            pmf_table = pmf_table / ((rowsum - knum - jnum) * (colsum(2) - jnum))
         end Until (pmf_table * pmfh > pmf_Obs) || (jnum < mode)
         pmf_table_jnum = pmf_table
      else
         Do
            pmf_table_jnum = pmf_table
            pmf_table = pmf_table * ((rowsum - knum - jnum) * (colsum(2) - jnum))
            jnum = jnum + 1.0
            pmf_table = pmf_table / (jnum * (col1mRowSum + knum + jnum))
            if pmf_table = 0.0
               ccdf = 0.0
               temp = 0.0
               Exit Do
            end
            if pmf_table * pmfh <= pmf_Obs Exit Do
            temp = temp - pmf_table
         end
         jnum = jnum - 1.0
      end
      if k = 50 pmf_table_jnum = pmf_hypergeometric(jnum, rowsum - knum, colsum(2), colsum12)
      ccdf = ccdf + temp
      if inum > jnum Exit Do
      Call AddValueToStack(ast, pmfh * (cdf + ccdf))
#Debug.Print knum, pmfh * (cdf + ccdf)
   end
   if pmfh > 0.0
      prob = cdf_hypergeometric(knum, rowsum, colsum(3), cs)
   end
   inum = inumstart
   jnum = jnumstart
   cdf = cdf_save
   ccdf = ccdf_save
   pmf_table_inum = pmf_table_inum_save
   pmf_table_jnum = pmf_table_jnum_save
   knum = ml(3)
   pmfh = pmfh_save
   k = 0
   while knum <= colsum(3)
      k = k + 1
      if k = 100
         knum = knum + 1.0
         pmfh = pmf_hypergeometric(knum, rowsum, colsum(3), cs)
         k = 0
      else
         pmfh = pmfh * ((rowsum - knum) * (colsum(3) - knum))
         knum = knum + 1.0
         pmfh = pmfh / (knum * (colsum12 - rowsum + knum))
         #pmfh = pmf_hypergeometric(knum, rowsum, colsum(3), cs)
      end
      if pmfh <= pmf_Obs Exit Do
      mode = Int((rowsum - knum + 1.0) * (colsum(2) + 1.0) / (colsum12 + 2.0))
      pmf_table = pmf_table_inum * ((rowsum - knum - inum + 1.0) * (colsum12 - rowsum + knum)) / ((rowsum - knum + 1.0) * (col1mRowSum + knum + inum))   #pmf_hypergeometric(inum, rowsum - knum, colsum(2), colsum12)
      temp = pmf_table_inum * inum / (rowsum - knum + 1.0)
      if pmf_table = 0.0
         inum = inum - 1.0
         pmf_table = temp * (colsum12 - rowsum + knum) / (colsum(2) - inum) #pmf_hypergeometric(inum, rowsum - knum, colsum(2), colsum12)
         temp = temp - pmf_table
      end
      while inum > mode
         pmf_table = pmf_table * (inum * (col1mRowSum + knum + inum))
         inum = inum - 1.0
         pmf_table = pmf_table / ((rowsum - knum - inum) * (colsum(2) - inum))
         temp = temp - pmf_table
      end
      if pmf_table * pmfh <= pmf_Obs
         Do
            temp = temp + pmf_table
            pmf_table = pmf_table * ((rowsum - knum - inum) * (colsum(2) - inum))
            inum = inum + 1.0
            pmf_table = pmf_table / (inum * (col1mRowSum + knum + inum))
         end Until (pmf_table * pmfh > pmf_Obs) || (inum > mode)
         pmf_table_inum = pmf_table
      else #if cdf > 0.0
         Do
            pmf_table_inum = pmf_table
            pmf_table = pmf_table * (inum * (col1mRowSum + knum + inum))
            inum = inum - 1.0
            pmf_table = pmf_table / ((rowsum - knum - inum) * (colsum(2) - inum))
            if pmf_table = 0.0
               cdf = 0.0
               temp = 0.0
               Exit Do
            end
            if pmf_table * pmfh <= pmf_Obs Exit Do
            temp = temp - pmf_table
         end
         inum = inum + 1.0
      end
      if k = 50 pmf_table_inum = pmf_hypergeometric(inum, rowsum - knum, colsum(2), colsum12)
      cdf = cdf + temp
      jnum = jnum - 1.0
      pmf_table = pmf_table_jnum * ((jnum + 1.0) * (colsum12 - rowsum + knum)) / ((rowsum - knum + 1.0) * (colsum(2) - jnum))    #pmf_hypergeometric(jnum, rowsum - knum, colsum(2), colsum12)
      temp = pmf_table_jnum * (rowsum - knum - jnum) / (rowsum - knum + 1.0)
      if pmf_table = 0.0
         jnum = jnum + 1.0
         pmf_table = temp * (colsum(1) - rowsum + knum - jnum) / (colsum12 - rowsum + knum)  #pmf_hypergeometric(jnum, rowsum - knum, colsum(2), colsum12)
         temp = temp - pmf_table
      end
      while jnum < mode # && pmf_table > 0.0
         pmf_table = pmf_table * ((rowsum - knum - jnum) * (colsum(2) - jnum))
         jnum = jnum + 1.0
         pmf_table = pmf_table / (jnum * (col1mRowSum + knum + jnum))
         temp = temp - pmf_table
      end
      if pmf_table * pmfh <= pmf_Obs
         Do
            temp = temp + pmf_table
            pmf_table = pmf_table * (jnum * (col1mRowSum + knum + jnum))
            jnum = jnum - 1.0
            pmf_table = pmf_table / ((rowsum - knum - jnum) * (colsum(2) - jnum))
         end Until (pmf_table * pmfh > pmf_Obs) || (jnum < mode)
         pmf_table_jnum = pmf_table
      else #if ccdf > 0.0
         Do
            pmf_table_jnum = pmf_table
            pmf_table = pmf_table * ((rowsum - knum - jnum) * (colsum(2) - jnum))
            jnum = jnum + 1.0
            pmf_table = pmf_table / (jnum * (col1mRowSum + knum + jnum))
            if pmf_table = 0.0
               ccdf = 0.0
               temp = 0.0
               Exit Do
            end
            if pmf_table * pmfh <= pmf_Obs Exit Do
            temp = temp - pmf_table
         end
         jnum = jnum - 1.0
      end
      if k = 50 pmf_table_jnum = pmf_hypergeometric(jnum, rowsum - knum, colsum(2), colsum12)
      ccdf = ccdf + temp
      if inum > jnum Exit Do
      Call AddValueToStack(ast, pmfh * (cdf + ccdf))
#Debug.Print knum, pmfh * (cdf + ccdf)
   end
   if pmfh > 0.0
      prob = prob + comp_cdf_hypergeometric(knum - 1.0, rowsum, colsum(3), cs)
   end
   fet_23 = prob + StackTotal(ast)
#Call DumpAddStack(ast)
#Debug.Print fet_23
else
   prob = 0.0
   Call InitAddStack(ast)
   inum = inumstart
   jnum = jnumstart
   prob_d = fet_23(c - 1, colsum, rowsum - knum, pmf_Obs / pmfh, inum, jnum)
#Debug.Print knum, pmfh, prob_d, pmfh * prob_d
   Call AddValueToStack(ast, pmfh * prob_d)
   inumstart = inum
   jnumstart = jnum
   if inum > jnum
      fet_23 = 1.0
      Exit Function
   end
   Do
      pmfh = pmfh * ((colsum(c) - knum) * (rowsum - knum))
      knum = knum + 1.0
      pmfh = pmfh / (knum * (cs - colsum(c) - rowsum + knum))
      if pmfh = 0.0 Exit Do
      prob_d = fet_23(c - 1, colsum, rowsum - knum, pmf_Obs / pmfh, inum, jnum)
      if inum > jnum
         prob = prob + comp_cdf_hypergeometric(knum - 1.0, colsum(c), rowsum, cs)
         Exit Do
      end
      Call AddValueToStack(ast, pmfh * prob_d)
   end
   pmfh = pmfh_save
   knum = ml(c)
   inum = inumstart
   jnum = jnumstart
   Do
      pmfh = pmfh * (knum * (cs - colsum(c) - rowsum + knum))
      knum = knum - 1.0
      pmfh = pmfh / ((colsum(c) - knum) * (rowsum - knum))
      if pmfh = 0.0 Exit Do
      prob_d = fet_23(c - 1, colsum, rowsum - knum, pmf_Obs / pmfh, inum, jnum)
      if inum > jnum
         prob = prob + cdf_hypergeometric(knum, colsum(c), rowsum, cs)
         Exit Do
      end
      Call AddValueToStack(ast, pmfh * prob_d)
   end
   fet_23 = prob + StackTotal(ast)
end
end

Function fet_24(c As Long, ByRef colsum()::Float64, rowsum::Float64, pmf_Obs::Float64, ByRef inumstart::Float64, ByRef jnumstart::Float64)::Float64
Dim d::Float64, cs::Float64, colsum12::Float64, colsum34::Float64, prob::Float64, temp::Float64, inum::Float64, jnum::Float64, inum_save::Float64, jnum_save::Float64
Dim cdf::Float64, ccdf::Float64, pmf_table_inum::Float64, pmf_table_jnum::Float64, mode::Float64, cdf_save::Float64, ccdf_save::Float64, col1mRowSum::Float64, probf4::Float64
Dim dnum_old::Float64, fnum_old::Float64, pmfd_old::Float64, dnum_save::Float64, pmfd_save::Float64
Dim i As Long, j As Long, count As Long
Dim all_d_zero::Bool
Dim asto As TAddStack

ReDim ml(1 To c)::Float64
if pmf_Obs > 1.0 #All tables have pmf <= 1
   fet_24 = 1.0
   Exit Function
end

cs = 0.0
For i = 1 To c
   cs = cs + colsum(i)
Next i
#First guess at mode vector
ml(1) = rowsum
For i = 2 To c
   ml(i) = Int(rowsum * colsum(i) / cs + 0.5)
   ml(1) = ml(1) - ml(i)
Next i

Do #Update guess at mode vector
   all_d_zero = true
   For i = 1 To c - 1
      For j = i + 1 To c
         d = ml(i) - Int((colsum(i) + 1.0) * (ml(i) + ml(j) + 1.0) / (colsum(i) + colsum(j) + 2.0))
         if d <> 0.0
            ml(i) = ml(i) - d
            ml(j) = ml(j) + d
            all_d_zero = false
         end
      Next j
   Next i
end Until all_d_zero

Dim pmff::Float64, pmfd::Float64, pmfd_down::Float64, pmfd_up::Float64
Dim cdff::Float64, probf::Float64, pmf_Obs_save::Float64
Dim dnum::Float64, dnum_up::Float64, dnum_down::Float64, fnum::Float64, fnum_save::Float64, pmff_save::Float64, rowsummfnum::Float64
Dim cdf_start::Float64, ccdf_start::Float64, inum_min::Float64, jnum_max::Float64
Dim continue_up::Bool, continue_down::Bool, exit_loop::Bool
if c = 4
   Call InitAddStack(asto)
   prob = 0.0
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
   inum_min = max(0.0, -(fnum + col1mRowSum))
   inum = min(max(inum_min, inumstart), ml(2))
   jnum_max = min(rowsummfnum, colsum(2))
   jnum = max(min(jnum_max, jnumstart), ml(2))
   pmf_table_inum = pmf_hypergeometric(inum, rowsummfnum, colsum(2), colsum12)
   mode = Int((rowsummfnum + 1.0) * (colsum(2) + 1.0) / (colsum12 + 2.0))
   while pmf_table_inum = 0.0
      inum = Int((inum + max(2, mode)) * 0.5)
      pmf_table_inum = pmf_hypergeometric(inum, rowsummfnum, colsum(2), colsum12)
   end
   while (pmf_table_inum > pmf_Obs) && (inum > inum_min)
      pmf_table_inum = pmf_table_inum * (inum * (col1mRowSum + fnum + inum))
      inum = inum - 1.0
      pmf_table_inum = pmf_table_inum / ((rowsummfnum - inum) * (colsum(2) - inum))
   end
   while (pmf_table_inum <= pmf_Obs) && (inum <= mode)
      pmf_table_inum = pmf_table_inum * (rowsummfnum - inum) * (colsum(2) - inum)
      inum = inum + 1.0
      pmf_table_inum = pmf_table_inum / (inum * (col1mRowSum + fnum + inum))
   end
   pmf_table_jnum = pmf_hypergeometric(jnum, rowsummfnum, colsum(2), colsum12)
   while pmf_table_jnum = 0.0
      jnum = Int((jnum + mode) * 0.5)
      pmf_table_jnum = pmf_hypergeometric(jnum, rowsummfnum, colsum(2), colsum12)
   end
   while (pmf_table_jnum > pmf_Obs) && (jnum < jnum_max)
      pmf_table_jnum = pmf_table_jnum * ((rowsummfnum - jnum) * (colsum(2) - jnum))
      jnum = jnum + 1.0
      pmf_table_jnum = pmf_table_jnum / (jnum * (col1mRowSum + fnum + jnum))
   end
   while (pmf_table_jnum < pmf_Obs) && (jnum >= mode)
      pmf_table_jnum = pmf_table_jnum * (jnum * (col1mRowSum + fnum + jnum))
      jnum = jnum - 1.0
      pmf_table_jnum = pmf_table_jnum / ((rowsummfnum - jnum) * (colsum(2) - jnum))
   end
   inumstart = inum
   jnumstart = jnum
   if inum > jnum
      fet_24 = 1.0
      Exit Function
   end
   cdf = cdf_hypergeometric(inum - 1, rowsummfnum, colsum(2), colsum12)
   ccdf = comp_cdf_hypergeometric(jnum, rowsummfnum, colsum(2), colsum12)
   cdf_start = cdf
   ccdf_start = ccdf
   continue_up = true
   continue_down = true
   inum_save = inum
   jnum_save = jnum
   cdf_save = cdf
   ccdf_save = ccdf
   cdf_save = cdf
   ccdf_save = ccdf
   prob = 0.0
   dnum_save = dnum
   pmfd_save = pmfd
   Do
      dnum_old = dnum
      fnum_old = fnum
      pmfd_old = pmfd
      probf4 = 0.0
      count = 1
      probf = pmfd * (cdf + ccdf)
      pmf_Obs = pmf_Obs * pmfd # was pmf_Obs = pmf_Obs / (pmfd * pmff)
      pmfd_down = pmfd
      pmfd_up = pmfd
      dnum_down = dnum
      dnum_up = dnum
      pmfd_down = pmfd_down * (dnum_down * (colsum(3) - fnum + dnum_down))
      dnum_down = dnum_down - 1
      pmfd_down = pmfd_down / ((fnum - dnum_down) * (colsum(4) - dnum_down))
      pmfd_up = pmfd_up * ((fnum - dnum_up) * (colsum(4) - dnum_up))
      dnum_up = dnum_up + 1.0
      pmfd_up = pmfd_up / (dnum_up * (colsum(3) - fnum + dnum_up))
      Do
         pmfd = max(pmfd_down, pmfd_up)
         if pmfd = 0.0 Exit Do
         while pmfd * pmf_table_inum <= pmf_Obs
            cdf = cdf + pmf_table_inum
            pmf_table_inum = pmf_table_inum * (rowsummfnum - inum) * (colsum(2) - inum)
            inum = inum + 1.0
            pmf_table_inum = pmf_table_inum / (inum * (col1mRowSum + fnum + inum))
            if (inum > mode) Exit Do
         end
         while pmfd * pmf_table_jnum <= pmf_Obs
            ccdf = ccdf + pmf_table_jnum
            pmf_table_jnum = pmf_table_jnum * (jnum * (col1mRowSum + fnum + jnum))
            jnum = jnum - 1.0
            pmf_table_jnum = pmf_table_jnum / ((rowsummfnum - jnum) * (colsum(2) - jnum))
            if (jnum < mode) Exit Do
         end
         if inum > jnum Exit Do
         probf4 = probf4 + (pmfd * (cdf + ccdf) - probf)
         count = count + 1
         if pmfd_down > pmfd_up
            pmfd_down = pmfd_down * (dnum_down * (colsum(3) - fnum + dnum_down))
            dnum_down = dnum_down - 1.0
            pmfd_down = pmfd_down / ((fnum - dnum_down) * (colsum(4) - dnum_down))
         else
            pmfd_up = pmfd_up * ((fnum - dnum_up) * (colsum(4) - dnum_up))
            dnum_up = dnum_up + 1.0
            pmfd_up = pmfd_up / (dnum_up * (colsum(3) - fnum + dnum_up))
         end
#Debug.Print inum, jnum, pmf_table_inum, pmf_table_jnum, cdf, ccdf, probf
      end
      probf = count * probf + probf4
      if inum > jnum
         pmfd_down = cdf_hypergeometric(dnum_down, fnum, colsum(4), colsum34)
         pmfd_up = comp_cdf_hypergeometric(dnum_up - 1.0, fnum, colsum(4), colsum34)
         probf = probf + pmfd_down + pmfd_up
      end

      Call AddValueToStack(asto, probf * pmff)
      if continue_up
         pmff = pmff * ((rowsummfnum) * (colsum34 - fnum))
         fnum = fnum + 1.0
         rowsummfnum = rowsum - fnum
         pmff = pmff / ((colsum12 - rowsummfnum) * (fnum))
         continue_up = pmff > 0.0
         if Not continue_up
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
         else
            inum = inum_save
            jnum = jnum_save
            cdf = cdf_save
            ccdf = ccdf_save
            inum = inum - 1.0
            temp = PBB(col1mRowSum + fnum + inum, rowsummfnum - inum, colsum(2) - inum, inum + 1.0)
            hTerm = hTerm * ((rowsummfnum - inum) * (colsum(2) - inum) * (colsum12 + 1.0))
            inum = inum + 1.0
            hTerm = hTerm / ((rowsummfnum + 1.0) * (colsum(2) + 1.0) * (col1mRowSum + fnum + inum))
            pmf_table_inum = hTerm #pmf_hypergeometric(inum, rowsummfnum, colsum(2), colsum12)
            cdf = cdf + temp
            temp = PBB(jnum, colsum(2) - jnum, rowsummfnum - jnum + 1.0, col1mRowSum + fnum + jnum)
            hTerm = hTerm * (jnum * (col1mRowSum + fnum + jnum) * (colsum12 + 1.0))
            jnum = jnum - 1.0
            hTerm = hTerm / ((rowsummfnum + 1.0) * (colsum(2) - jnum) * (colsum(1) + 1.0))
            pmf_table_jnum = hTerm #pmf_hypergeometric(jnum, rowsummfnum, colsum(2), colsum12)
            ccdf = ccdf + temp
         end
      end
      Do
         if Not continue_up
            pmff = pmff * ((colsum12 - rowsummfnum) * (fnum))
            fnum = fnum - 1.0
            rowsummfnum = rowsum - fnum
            pmff = pmff / ((rowsummfnum) * (colsum34 - fnum))
            continue_down = pmff > 0.0
            if Not continue_down Exit Do
            inum = inum_save
            jnum = jnum_save
            cdf = cdf_save
            ccdf = ccdf_save
            temp = PBB(inum, colsum(2) - inum, rowsummfnum - inum, col1mRowSum + fnum + inum + 1.0)
            hTerm = hTerm * ((rowsummfnum - inum) * (colsum(2) - inum) * (colsum12 + 1.0))
            inum = inum + 1.0
            hTerm = hTerm / (inum * (colsum(1) + 1.0) * (colsum12 - rowsummfnum + 1.0))
            pmf_table_inum = hTerm #pmf_hypergeometric(inum, rowsummfnum, colsum(2), colsum12)
            cdf = cdf + temp
            temp = PBB(col1mRowSum + fnum + jnum + 1.0, rowsummfnum - jnum - 1.0, colsum(2) - jnum, jnum + 1.0)
            hTerm = hTerm * ((jnum + 1.0) * (col1mRowSum + fnum + jnum + 1.0) * (colsum12 + 1.0)) / ((colsum(2) + 1.0) * (rowsummfnum - jnum) * (colsum12 - rowsummfnum + 1.0))
            pmf_table_jnum = hTerm #pmf_hypergeometric(jnum, rowsummfnum, colsum(2), colsum12)
            ccdf = ccdf + temp
         end
#Initialise for next go round loop
         mode = Int((rowsummfnum + 1.0) * (colsum(2) + 1.0) / (colsum12 + 2.0))
         dnum = Int((fnum + 1.0) * (colsum(4) + 1.0) / (colsum34 + 2.0))
         if continue_up
            if dnum = dnum_old
               pmfd = pmfd_old * ((colsum34 - fnum_old - colsum(4) + dnum) * (fnum)) / ((fnum - dnum) * (colsum34 - fnum_old))
            elseif dnum = dnum_old + 1.0
               pmfd = pmfd_old * (fnum * (colsum(4) - dnum_old)) / (dnum * (colsum34 - fnum_old))
            else
               Debug.Print dnum_old, fnum_old, dnum, fnum
               pmfd = 1.0 / 0.0
               pmfd = pmf_hypergeometric(dnum, fnum, colsum(4), colsum34)
            end
         else
            if dnum = dnum_old
               pmfd = pmfd_old * ((fnum_old - dnum) * (colsum34 - fnum)) / ((colsum34 - fnum - colsum(4) + dnum) * (fnum_old))
            elseif dnum = dnum_old - 1.0
               pmfd = pmfd_old * (dnum_old * (colsum34 - fnum)) / (fnum_old * (colsum(4) - dnum))
            else
               Debug.Print dnum_old, fnum_old, dnum, fnum
               pmfd = 1.0 / 0.0
               pmfd = pmf_hypergeometric(dnum, fnum, colsum(4), colsum34)
            end
         end
         #pmfd = pmf_hypergeometric(dnum, fnum, colsum(4), colsum34)
         pmf_Obs = pmf_Obs_save / (pmff * pmfd)
         inum_min = max(0.0, -(fnum + col1mRowSum))
         jnum_max = min(rowsummfnum, colsum(2))
         if jnum_max >= 0.0 && pmf_Obs < 1.0
            while (pmf_table_inum > pmf_Obs) && (inum > inum_min) || (inum > mode)
               pmf_table_inum = pmf_table_inum * (inum * (col1mRowSum + fnum + inum))
               inum = inum - 1.0
               pmf_table_inum = pmf_table_inum / ((rowsummfnum - inum) * (colsum(2) - inum))
               cdf = cdf - pmf_table_inum
            end
            while (pmf_table_inum < pmf_Obs) && (inum <= mode)
               cdf = cdf + pmf_table_inum
               pmf_table_inum = pmf_table_inum * (rowsummfnum - inum) * (colsum(2) - inum)
               inum = inum + 1.0
               pmf_table_inum = pmf_table_inum / (inum * (col1mRowSum + fnum + inum))
            end
            if inum = inum_min cdf = 0.0
            while (pmf_table_jnum > pmf_Obs) && (jnum < jnum_max) || (jnum < mode)
               pmf_table_jnum = pmf_table_jnum * ((rowsummfnum - jnum) * (colsum(2) - jnum))
               jnum = jnum + 1.0
               pmf_table_jnum = pmf_table_jnum / (jnum * (col1mRowSum + fnum + jnum))
               ccdf = ccdf - pmf_table_jnum
            end
            while (pmf_table_jnum < pmf_Obs) && (jnum >= mode)
               ccdf = ccdf + pmf_table_jnum
               pmf_table_jnum = pmf_table_jnum * (jnum * (col1mRowSum + fnum + jnum))
               jnum = jnum - 1.0
               pmf_table_jnum = pmf_table_jnum / ((rowsummfnum - jnum) * (colsum(2) - jnum))
            end
            if jnum = jnum_max ccdf = 0.0
         else
            inum = jnum + 1.0
         end
         exit_loop = true
         if inum > jnum
            if continue_up
               prob = prob + comp_cdf_hypergeometric(fnum - 1.0, colsum34, rowsum, cs)
               continue_up = false
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
               exit_loop = false
            else
               prob = prob + cdf_hypergeometric(fnum, colsum34, rowsum, cs)
               continue_down = false
            end
         end
      end Until exit_loop
      if Not continue_down Exit Do
      cdf_save = cdf
      ccdf_save = ccdf
      inum_save = inum
      jnum_save = jnum
   end Until !continue_down
   fet_24 = prob + StackTotal(asto)
elseif c > 4
   Dim knum::Float64
   knum = ml(c)
   pmff = pmf_hypergeometric(knum, colsum(c), rowsum, cs)
   pmff_save = pmff
   inum = inumstart
   jnum = jnumstart
   Call InitAddStack(asto)
   probf4 = fet_24(c - 1, colsum, rowsum - knum, pmf_Obs / pmff, inum, jnum)
   inumstart = inum
   jnumstart = jnum
   if inum > jnum
      fet_24 = 1.0
      Exit Function
   end
   Call AddValueToStack(asto, pmff * probf4)

   Do
      pmff = pmff * ((colsum(c) - knum) * (rowsum - knum))
      knum = knum + 1.0
      pmff = pmff / (knum * (cs - colsum(c) - rowsum + knum))
      if pmff = 0.0 Exit Do
      probf4 = fet_24(c - 1, colsum, rowsum - knum, pmf_Obs / pmff, inum, jnum)
      if inum > jnum
         prob = prob + comp_cdf_hypergeometric(knum - 1.0, colsum(c), rowsum, cs)
         Exit Do
      end
      Call AddValueToStack(asto, pmff * probf4)
   end
   pmff = pmff_save
   knum = ml(c)
   inum = inumstart
   jnum = jnumstart
   Do
      pmff = pmff * (knum * (cs - colsum(c) - rowsum + knum))
      knum = knum - 1.0
      pmff = pmff / ((colsum(c) - knum) * (rowsum - knum))
      if pmff = 0.0 Exit Do
      probf4 = fet_24(c - 1, colsum, rowsum - knum, pmf_Obs / pmff, inum, jnum)
      if inum > jnum
         prob = prob + cdf_hypergeometric(knum, colsum(c), rowsum, cs)
         Exit Do
      end
      Call AddValueToStack(asto, pmff * probf4)
   end
   fet_24 = prob + StackTotal(asto)
else
   fet_24 = "Values other than 4 not allowed"
end
end

Function fet_25(c As Long, ByRef colsum()::Float64, rowsum::Float64, pmf_Obs::Float64, ByRef inumstart::Float64, ByRef jnumstart::Float64)::Float64
Dim d::Float64, cs::Float64, prob::Float64, temp::Float64, inum::Float64, jnum::Float64, inum_save::Float64, jnum_save::Float64, inum_save_save::Float64, jnum_save_save::Float64
Dim cdf::Float64, ccdf::Float64, pmf_table_inum::Float64, pmf_table_jnum::Float64, mode::Float64, cdf_save::Float64, ccdf_save::Float64, col1mRowSum::Float64, prob_d::Float64
Dim i As Long, j As Long, c_count As Long
Dim all_d_zero::Bool

ReDim ml(1 To c)::Float64

if pmf_Obs > 1.0 #All tables have pmf <= 1
   fet_25 = 1.0
   Exit Function
end

cs = 0.0
For i = 1 To c
   cs = cs + colsum(i)
Next i
#First guess at mode vector
ml(1) = rowsum
For i = 2 To c
   ml(i) = Int(rowsum * colsum(i) / cs + 0.5)
   ml(1) = ml(1) - ml(i)
Next i

Do #Update guess at mode vector
   all_d_zero = true
   For i = 1 To c - 1
      For j = i + 1 To c
         d = ml(i) - Int((colsum(i) + 1.0) * (ml(i) + ml(j) + 1.0) / (colsum(i) + colsum(j) + 2.0))
         if d <> 0.0
            ml(i) = ml(i) - d
            ml(j) = ml(j) + d
            all_d_zero = false
         end
      Next j
   Next i
end Until all_d_zero

Dim pmff::Float64, pmfd4::Float64, pmfd4_down::Float64, pmfd4_up::Float64, pmfd5::Float64, pmfd5_save::Float64
Dim cdff::Float64, probf::Float64, pmf_Obs_save::Float64, probf5::Float64, cdf_save_save::Float64, ccdf_save_save::Float64
Dim d4num::Float64, d4num_up::Float64, d4num_down::Float64, d5num::Float64, d5num_save::Float64, fnum::Float64, fnum_save::Float64, pmff_save::Float64, rowsummfnum::Float64
Dim cdf_start::Float64, ccdf_start::Float64, inum_min::Float64, jnum_max::Float64, pmf_table_inum_save::Float64, pmf_table_jnum_save::Float64
Dim inum_save5::Float64, jnum_save5::Float64, pmf_table_inum_save5::Float64, pmf_table_jnum_save5::Float64, cdf_save5::Float64, ccdf_save5::Float64
Dim colsum12::Float64, colsum34::Float64, colsum345::Float64

Dim continue_up::Bool, continue_down::Bool, exit_loop::Bool
Dim c5_up::Bool, c5_down::Bool, el5::Bool
if c = 5
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
   inum_min = max(0.0, -(fnum + col1mRowSum))
   inum = min(max(inum_min, inumstart), ml(2))
   jnum_max = min(rowsummfnum, colsum(2))
   jnum = max(min(jnum_max, jnumstart), ml(2))
   pmf_table_inum = pmf_hypergeometric(inum, rowsummfnum, colsum(2), colsum12)
   mode = Int((rowsummfnum + 1.0) * (colsum(2) + 1.0) / (colsum12 + 2.0))
   while pmf_table_inum = 0.0
      inum = Int((inum + max(2, mode)) * 0.5)
      pmf_table_inum = pmf_hypergeometric(inum, rowsummfnum, colsum(2), colsum12)
   end
   while (pmf_table_inum > pmf_Obs) && (inum > inum_min)
      pmf_table_inum = pmf_table_inum * (inum * (col1mRowSum + fnum + inum))
      inum = inum - 1.0
      pmf_table_inum = pmf_table_inum / ((rowsummfnum - inum) * (colsum(2) - inum))
   end
   while (pmf_table_inum <= pmf_Obs) && (inum <= mode)
      pmf_table_inum = pmf_table_inum * (rowsummfnum - inum) * (colsum(2) - inum)
      inum = inum + 1.0
      pmf_table_inum = pmf_table_inum / (inum * (col1mRowSum + fnum + inum))
   end
   pmf_table_jnum = pmf_hypergeometric(jnum, rowsummfnum, colsum(2), colsum12)
   while pmf_table_jnum = 0.0
      jnum = Int((jnum + mode) * 0.5)
      pmf_table_jnum = pmf_hypergeometric(jnum, rowsummfnum, colsum(2), colsum12)
   end
   while (pmf_table_jnum > pmf_Obs) && (jnum < jnum_max)
      pmf_table_jnum = pmf_table_jnum * ((rowsummfnum - jnum) * (colsum(2) - jnum))
      jnum = jnum + 1.0
      pmf_table_jnum = pmf_table_jnum / (jnum * (col1mRowSum + fnum + jnum))
   end
   while (pmf_table_jnum < pmf_Obs) && (jnum >= mode)
      pmf_table_jnum = pmf_table_jnum * (jnum * (col1mRowSum + fnum + jnum))
      jnum = jnum - 1.0
      pmf_table_jnum = pmf_table_jnum / ((rowsummfnum - jnum) * (colsum(2) - jnum))
   end
   inumstart = inum
   jnumstart = jnum
   if inum > jnum
      fet_25 = 1.0
      Exit Function
   end
   cdf = cdf_hypergeometric(inum - 1, rowsummfnum, colsum(2), colsum12)
   ccdf = comp_cdf_hypergeometric(jnum, rowsummfnum, colsum(2), colsum12)
   cdf_start = cdf
   ccdf_start = ccdf
   continue_up = true
   continue_down = true
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
   prob = 0.0
   probf5 = 0.0
   Do
      d5num_save = d5num
      pmfd5_save = pmfd5
      inum_save5 = inum
      jnum_save5 = jnum
      pmf_table_inum_save5 = pmf_table_inum
      pmf_table_jnum_save5 = pmf_table_jnum
      cdf_save5 = cdf
      ccdf_save5 = ccdf
      c5_up = true
      c5_down = true
      Do
         probf = pmfd4 * (cdf + ccdf)
         pmf_Obs = pmf_Obs * pmfd4 # was pmf_Obs = pmf_Obs / (pmfd4 * pmfd5 * pmff)
         pmfd4_down = pmfd4
         pmfd4_up = pmfd4
         d4num_down = d4num
         d4num_up = d4num
         pmfd4_down = pmfd4_down * d4num_down * (colsum(3) - fnum + d5num + d4num_down)
         d4num_down = d4num_down - 1
         pmfd4_down = pmfd4_down / ((fnum - d5num - d4num_down) * (colsum(4) - d4num_down))
         pmfd4_up = pmfd4_up * ((fnum - d5num - d4num_up) * (colsum(4) - d4num_up))
         d4num_up = d4num_up + 1.0
         pmfd4_up = pmfd4_up / (d4num_up * (colsum(3) - fnum + d5num + d4num_up))
         Do
            pmfd4 = max(pmfd4_down, pmfd4_up)
            if pmfd4 = 0.0 Exit Do
            if pmfd4 * pmf_table_inum <= pmf_Obs
               Do
                  cdf = cdf + pmf_table_inum
                  pmf_table_inum = pmf_table_inum * (rowsummfnum - inum) * (colsum(2) - inum)
                  inum = inum + 1.0
                  pmf_table_inum = pmf_table_inum / (inum * (col1mRowSum + fnum + inum))
               end Until (pmfd4 * pmf_table_inum > pmf_Obs) || (inum > mode)
            end
            if pmfd4 * pmf_table_jnum < pmf_Obs
               Do
                  ccdf = ccdf + pmf_table_jnum
                  pmf_table_jnum = pmf_table_jnum * (jnum * (col1mRowSum + fnum + jnum))
                  jnum = jnum - 1.0
                  pmf_table_jnum = pmf_table_jnum / ((rowsummfnum - jnum) * (colsum(2) - jnum))
               end Until (pmfd4 * pmf_table_jnum > pmf_Obs) || (jnum < mode)
            end
            if inum > jnum
               pmfd4_down = cdf_hypergeometric(d4num_down, fnum - d5num, colsum(4), colsum34)
               pmfd4_up = comp_cdf_hypergeometric(d4num_up - 1.0, fnum - d5num, colsum(4), colsum34)
               probf = probf + pmfd4_down + pmfd4_up
               Exit Do
            end
            probf = probf + pmfd4 * (cdf + ccdf)
            if pmfd4_down > pmfd4_up
               pmfd4_down = pmfd4_down * d4num_down * (colsum(3) - fnum + d5num + d4num_down)
               d4num_down = d4num_down - 1.0
               pmfd4_down = pmfd4_down / ((fnum - d5num - d4num_down) * (colsum(4) - d4num_down))
            else
               pmfd4_up = pmfd4_up * ((fnum - d5num - d4num_up) * (colsum(4) - d4num_up))
               d4num_up = d4num_up + 1.0
               pmfd4_up = pmfd4_up / (d4num_up * (colsum(3) - fnum + d5num + d4num_up))
            end
#Debug.Print inum, jnum, pmf_table_inum, pmf_table_jnum, cdf, ccdf, probf
         end
         probf5 = probf5 + probf * pmfd5
         if c5_up
            pmfd5 = pmfd5 * ((fnum - d5num) * (colsum(5) - d5num))
            d5num = d5num + 1.0
            pmfd5 = pmfd5 / (d5num * (colsum34 - fnum + d5num))
            c5_up = pmfd5 > 0.0
            if Not c5_up
               pmfd5 = pmfd5_save
               d5num = d5num_save
               inum_save_save = inum_save5
               jnum_save_save = jnum_save5
               pmf_table_inum_save = pmf_table_inum_save5
               pmf_table_jnum_save = pmf_table_jnum_save5
               cdf_save_save = cdf_save5
               ccdf_save_save = ccdf_save5
            end
         end
         Do
            if Not c5_up
               pmfd5 = pmfd5 * d5num * (colsum34 - fnum + d5num)
               d5num = d5num - 1.0
               pmfd5 = pmfd5 / ((fnum - d5num) * (colsum(5) - d5num))
               c5_down = pmfd5 > 0.0
               if Not c5_down Exit Do
            end
            d4num = Int((fnum - d5num + 1.0) * (colsum(4) + 1.0) / (colsum34 + 2.0))
            pmfd4 = pmf_hypergeometric(d4num, fnum - d5num, colsum(4), colsum34)
            pmf_Obs = pmf_Obs_save / (pmfd4 * pmfd5 * pmff)
            inum = inum_save_save
            jnum = jnum_save_save
            pmf_table_inum = pmf_table_inum_save
            pmf_table_jnum = pmf_table_jnum_save
            cdf = cdf_save_save
            ccdf = ccdf_save_save
            #while (pmf_table_inum > pmf_Obs) && (inum > inum_min)
            #   pmf_table_inum = pmf_table_inum * (inum * (col1mRowSum + fnum + inum))
            #   inum = inum - 1.0
            #   pmf_table_inum = pmf_table_inum / ((rowsummfnum - inum) * (colsum(2) - inum))
            #   cdf = cdf - pmf_table_inum
            #end
            if pmf_table_inum * (inum * (col1mRowSum + fnum + inum)) / ((rowsummfnum - inum + 1.0) * (colsum(2) - inum + 1.0)) > pmf_Obs
               fet_25 = "Problem with cdf"
            end
            while (pmf_table_inum <= pmf_Obs) && (inum <= mode)
               cdf = cdf + pmf_table_inum
               pmf_table_inum = pmf_table_inum * (rowsummfnum - inum) * (colsum(2) - inum)
               inum = inum + 1.0
               pmf_table_inum = pmf_table_inum / (inum * (col1mRowSum + fnum + inum))
            end
            #while (pmf_table_jnum > pmf_Obs) && (jnum < jnum_max)
            #   pmf_table_jnum = pmf_table_jnum * ((rowsummfnum - jnum) * (colsum(2) - jnum))
            #   jnum = jnum + 1.0
            #   pmf_table_jnum = pmf_table_jnum / (jnum * (col1mRowSum + fnum + jnum))
            #   ccdf = ccdf - pmf_table_jnum
            #end
            if pmf_table_jnum * ((rowsummfnum - jnum) * (colsum(2) - jnum)) / ((jnum + 1.0) * (col1mRowSum + fnum + jnum + 1.0)) > pmf_Obs
               fet_25 = "Problem with ccdf"
            end
            while (pmf_table_jnum < pmf_Obs) && (jnum >= mode)
               ccdf = ccdf + pmf_table_jnum
               pmf_table_jnum = pmf_table_jnum * (jnum * (col1mRowSum + fnum + jnum))
               jnum = jnum - 1.0
               pmf_table_jnum = pmf_table_jnum / ((rowsummfnum - jnum) * (colsum(2) - jnum))
            end
            el5 = true
            if inum > jnum
               if c5_up
                  probf5 = probf5 + comp_cdf_hypergeometric(d5num - 1.0, fnum, colsum(5), colsum345)
                  pmfd5 = pmfd5_save
                  d5num = d5num_save
                  inum_save_save = inum_save5
                  jnum_save_save = jnum_save5
                  pmf_table_inum_save = pmf_table_inum_save5
                  pmf_table_jnum_save = pmf_table_jnum_save5
                  cdf_save_save = cdf_save5
                  ccdf_save_save = ccdf_save5
                  el5 = false
                  c5_up = false
               else
                  probf5 = probf5 + cdf_hypergeometric(d5num, fnum, colsum(5), colsum345)
                  c5_down = false
                  Exit Do
               end
            end
         end Until el5
         if Not c5_down Exit Do
         inum_save_save = inum
         jnum_save_save = jnum
         pmf_table_inum_save = pmf_table_inum
         pmf_table_jnum_save = pmf_table_jnum
         cdf_save_save = cdf
         ccdf_save_save = ccdf
      end
      prob = prob + probf5 * pmff
#Debug.Print prob
      probf5 = 0.0
      if continue_up
         pmff = pmff * ((rowsummfnum) * (colsum345 - fnum))
         fnum = fnum + 1.0
         rowsummfnum = rowsum - fnum
         pmff = pmff / ((colsum12 - rowsummfnum) * (fnum))
         continue_up = pmff > 0.0
         if Not continue_up
            pmff = pmff_save
            fnum = fnum_save
            rowsummfnum = rowsum - fnum
            inum_save = inumstart
            jnum_save = jnumstart
            cdf_save = cdf_start
            ccdf_save = ccdf_start
         else
            inum = inum_save
            jnum = jnum_save
            cdf = cdf_save
            ccdf = ccdf_save
            inum = inum - 1.0
            temp = PBB(col1mRowSum + fnum + inum, rowsummfnum - inum, colsum(2) - inum, inum + 1.0)
            hTerm = hTerm * ((rowsummfnum - inum) * (colsum(2) - inum) * (colsum12 + 1.0))
            inum = inum + 1.0
            hTerm = hTerm / ((rowsummfnum + 1.0) * (colsum(2) + 1.0) * (col1mRowSum + fnum + inum))
            pmf_table_inum = hTerm #pmf_hypergeometric(inum, rowsummfnum, colsum(2), colsum12)
            cdf = cdf + temp
            temp = PBB(jnum, colsum(2) - jnum, rowsummfnum - jnum + 1.0, col1mRowSum + fnum + jnum)
            hTerm = hTerm * (jnum * (col1mRowSum + fnum + jnum) * (colsum12 + 1.0))
            jnum = jnum - 1.0
            hTerm = hTerm / ((rowsummfnum + 1.0) * (colsum(2) - jnum) * (colsum(1) + 1.0))
            pmf_table_jnum = hTerm #pmf_hypergeometric(jnum, rowsummfnum, colsum(2), colsum12)
            ccdf = ccdf + temp
         end
      end
      Do
         if Not continue_up
            pmff = pmff * ((colsum12 - rowsummfnum) * (fnum))
            fnum = fnum - 1.0
            rowsummfnum = rowsum - fnum
            pmff = pmff / ((rowsummfnum) * (colsum345 - fnum))
            continue_down = pmff > 0.0
            if Not continue_down Exit Do
            inum = inum_save
            jnum = jnum_save
            cdf = cdf_save
            ccdf = ccdf_save
            temp = PBB(inum, colsum(2) - inum, rowsummfnum - inum, col1mRowSum + fnum + inum + 1.0)
            hTerm = hTerm * (rowsummfnum - inum) * (colsum(2) - inum) * (colsum12 + 1.0)
            inum = inum + 1.0
            hTerm = hTerm / (inum * (colsum(1) + 1.0) * (colsum12 - rowsummfnum + 1.0))
            pmf_table_inum = hTerm #pmf_hypergeometric(inum, rowsummfnum, colsum(2), colsum12)
            cdf = cdf + temp
            temp = PBB(col1mRowSum + fnum + jnum + 1.0, rowsummfnum - jnum - 1.0, colsum(2) - jnum, jnum + 1.0)
            hTerm = hTerm * ((jnum + 1.0) * (col1mRowSum + fnum + jnum + 1.0) * (colsum12 + 1.0)) / ((colsum(2) + 1.0) * (rowsummfnum - jnum) * (colsum12 - rowsummfnum + 1.0))
            pmf_table_jnum = hTerm #pmf_hypergeometric(jnum, rowsummfnum, colsum(2), colsum12)
            ccdf = ccdf + temp
         end
         mode = Int((rowsummfnum + 1.0) * (colsum(2) + 1.0) / (colsum12 + 2.0))
         ml(4) = Int(fnum * colsum(4) / (colsum345) + 0.5)
         ml(5) = Int(fnum * colsum(5) / (colsum345) + 0.5)
         ml(3) = fnum - ml(4) - ml(5)
         Do #Update guess at mode vector
            all_d_zero = true
            For i = 3 To c - 1
               For j = i + 1 To c
                  d = ml(i) - Int((colsum(i) + 1.0) * (ml(i) + ml(j) + 1.0) / (colsum(i) + colsum(j) + 2.0))
                  if d <> 0.0
                     ml(i) = ml(i) - d
                     ml(j) = ml(j) + d
                     all_d_zero = false
                  end
               Next j
            Next i
         end Until all_d_zero
#Debug.Print ml(3), ml(4), ml(5)
         d5num = ml(5)
         d4num = ml(4)
         #if ml(4) <> Int((fnum - d5num + 1.0) * (colsum(4) + 1.0) / (colsum34 + 2.0))
         #   fet_25 = "Problem with mode"
         #end
         pmfd4 = pmf_hypergeometric(d4num, fnum - d5num, colsum(4), colsum34)
         pmfd5 = pmf_hypergeometric(d5num, fnum, colsum(5), colsum345)
         pmf_Obs = pmf_Obs_save / (pmfd4 * pmfd5 * pmff)
         inum_min = max(0.0, -(fnum + col1mRowSum))
         jnum_max = min(rowsummfnum, colsum(2))
         if jnum_max > 0.0 && pmf_Obs < 1.0
            while (pmf_table_inum > pmf_Obs) && (inum > inum_min) || (inum > mode)
               pmf_table_inum = pmf_table_inum * (inum * (col1mRowSum + fnum + inum))
               inum = inum - 1.0
               pmf_table_inum = pmf_table_inum / ((rowsummfnum - inum) * (colsum(2) - inum))
               cdf = cdf - pmf_table_inum
            end
            while (pmf_table_inum < pmf_Obs) && (inum <= mode)
               cdf = cdf + pmf_table_inum
               pmf_table_inum = pmf_table_inum * (rowsummfnum - inum) * (colsum(2) - inum)
               inum = inum + 1.0
               pmf_table_inum = pmf_table_inum / (inum * (col1mRowSum + fnum + inum))
            end
            if inum = inum_min cdf = 0.0
            while (pmf_table_jnum > pmf_Obs) && (jnum < jnum_max) || (jnum < mode)
               pmf_table_jnum = pmf_table_jnum * ((rowsummfnum - jnum) * (colsum(2) - jnum))
               jnum = jnum + 1.0
               pmf_table_jnum = pmf_table_jnum / (jnum * (col1mRowSum + fnum + jnum))
               ccdf = ccdf - pmf_table_jnum
            end
            while (pmf_table_jnum < pmf_Obs) && (jnum >= mode)
               ccdf = ccdf + pmf_table_jnum
               pmf_table_jnum = pmf_table_jnum * (jnum * (col1mRowSum + fnum + jnum))
               jnum = jnum - 1.0
               pmf_table_jnum = pmf_table_jnum / ((rowsummfnum - jnum) * (colsum(2) - jnum))
            end
            if jnum = jnum_max ccdf = 0.0
         else
            inum = jnum + 1.0
         end
         exit_loop = true
         if inum > jnum
            if continue_up
               prob = prob + comp_cdf_hypergeometric(fnum - 1.0, colsum345, rowsum, cs)
#Debug.Print prob
               continue_up = false
               pmff = pmff_save
               fnum = fnum_save
               rowsummfnum = rowsum - fnum
               inum_save = inumstart
               jnum_save = jnumstart
               cdf_save = cdf_start
               ccdf_save = ccdf_start
               exit_loop = false
            else
               prob = prob + cdf_hypergeometric(fnum, colsum345, rowsum, cs)
#Debug.Print prob
               continue_down = false
            end
         end
      end Until exit_loop
      if Not continue_down Exit Do
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
   end Until !continue_down
   fet_25 = prob
   Exit Function
elseif c >= 6
   Dim knum::Float64
   knum = ml(c)
   pmff = pmf_hypergeometric(knum, colsum(c), rowsum, cs)
   pmff_save = pmff
   inum = inumstart
   jnum = jnumstart
   prob_d = fet_25(c - 1, colsum, rowsum - knum, pmf_Obs / pmff, inum, jnum)
   inumstart = inum
   jnumstart = jnum
   if inum > jnum
      fet_25 = 1.0
      Exit Function
   end
   prob = pmff * prob_d

   Do
      pmff = pmff * ((colsum(c) - knum) * (rowsum - knum))
      knum = knum + 1.0
      pmff = pmff / (knum * (cs - colsum(c) - rowsum + knum))
      if pmff = 0.0 Exit Do
      prob_d = fet_25(c - 1, colsum, rowsum - knum, pmf_Obs / pmff, inum, jnum)
      if inum > jnum
         prob = prob + comp_cdf_hypergeometric(knum - 1.0, colsum(c), rowsum, cs)
         Exit Do
      end
      prob = prob + pmff * prob_d
   end
   pmff = pmff_save
   knum = ml(c)
   inum = inumstart
   jnum = jnumstart
   Do
      pmff = pmff * (knum * (cs - colsum(c) - rowsum + knum))
      knum = knum - 1.0
      pmff = pmff / ((colsum(c) - knum) * (rowsum - knum))
      if pmff = 0.0 Exit Do
      prob_d = fet_25(c - 1, colsum, rowsum - knum, pmf_Obs / pmff, inum, jnum)
      if inum > jnum
         prob = prob + cdf_hypergeometric(knum, colsum(c), rowsum, cs)
         Exit Do
      end
      prob = prob + pmff * prob_d
   end
   fet_25 = prob
else
   fet_25 = "c must be >= 5"
end
end

Function fet_26(c As Long, ByRef colsum()::Float64, rowsum::Float64, pmf_Obs::Float64, ByRef inumstart::Float64, ByRef jnumstart::Float64)::Float64
Dim d::Float64, cs::Float64, colsum1_2::Float64, prob::Float64, temp::Float64, inum::Float64, jnum::Float64, inum_save::Float64, jnum_save::Float64
Dim cdf::Float64, ccdf::Float64, pmf_table_inum::Float64, pmf_table_jnum::Float64, mode::Float64, col1mRowSum::Float64, row3sum::Float64
Dim i As Long, j As Long, tc As Long
Dim all_d_zero::Bool

Dim d4num::Float64, d4num_down::Float64, d4num_up::Float64
Dim pmff::Float64, pmfd4::Float64, pmfd4_down::Float64, pmfd4_up::Float64
Dim cdff::Float64, probf::Float64, pmf_Obs_save::Float64, cdf_start::Float64, ccdf_start::Float64
Dim fnum::Float64, fnum_start::Float64, pmff_start::Float64, rowsummfnum::Float64
Dim cdf_save::Float64, ccdf_save::Float64, inum_min::Float64, jnum_max::Float64, pmf_table_inum_save::Float64, pmf_table_jnum_save::Float64

Dim continue_up::Bool, continue_down::Bool, exit_loop::Bool
Dim el5::Bool

ReDim ml(1 To c)::Float64, dnum(5 To c)::Float64, dnum_up(5 To c)::Float64, dnum_down(5 To c)::Float64, dnum_save(5 To c)::Float64, colsumsum(3 To c)::Float64
ReDim pmfd(5 To c)::Float64, pmfd_save(5 To c)::Float64, probf5(5 To c)::Float64
ReDim inum_save5(5 To c)::Float64, jnum_save5(5 To c)::Float64, pmf_table_inum_save5(5 To c)::Float64, pmf_table_jnum_save5(5 To c)::Float64, cdf_save5(5 To c)::Float64, ccdf_save5(5 To c)::Float64
ReDim c_up(5 To c)::Bool, c_down(5 To c)::Bool
ReDim dnumsum(4 To c)::Float64, pmf_prod(4 To c + 1)
ReDim inum_next(5 To c)::Float64, jnum_next(5 To c)::Float64, pmf_table_inum_next(5 To c)::Float64, pmf_table_jnum_next(5 To c)::Float64, cdf_next(5 To c)::Float64, ccdf_next(5 To c)::Float64

if pmf_Obs > 1.0 #All tables have pmf <= 1
   fet_26 = 1.0
   Exit Function
end
pmf_Obs_save = pmf_Obs

colsumsum(3) = colsum(3)
For i = 4 To c
   colsumsum(i) = colsumsum(i - 1) + colsum(i)
Next i
colsum1_2 = colsum(1) + colsum(2)
cs = colsum1_2 + colsumsum(c)

#First guess at mode vector
ml(1) = rowsum
For i = 2 To c
   ml(i) = Int(rowsum * colsum(i) / cs + 0.5)
   ml(1) = ml(1) - ml(i)
Next i

Do #Update guess at mode vector
   all_d_zero = true
   For i = 1 To c - 1
      For j = i + 1 To c
         d = ml(i) - Int((colsum(i) + 1.0) * (ml(i) + ml(j) + 1.0) / (colsum(i) + colsum(j) + 2.0))
         if d <> 0.0
            ml(i) = ml(i) - d
            ml(j) = ml(j) + d
            all_d_zero = false
         end
      Next j
   Next i
end Until all_d_zero

#if c > = 5 && colsum(6) > 1
if c >= 5
   d4num = ml(4)
   dnumsum(c) = 0.0
   For i = c To 5 Step -1
      probf5(i) = 0.0
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
   inum_min = max(0.0, -(fnum + col1mRowSum))
   inum = min(max(inum_min, inumstart), ml(2))
   jnum_max = min(rowsummfnum, colsum(2))
   jnum = max(min(jnum_max, jnumstart), ml(2))
   pmf_table_inum = pmf_hypergeometric(inum, rowsummfnum, colsum(2), colsum1_2)
   mode = Int((rowsummfnum + 1.0) * (colsum(2) + 1.0) / (colsum1_2 + 2.0))
   while pmf_table_inum = 0.0
      inum = Int((inum + max(2, mode)) * 0.5)
      pmf_table_inum = pmf_hypergeometric(inum, rowsummfnum, colsum(2), colsum1_2)
   end
   while (pmf_table_inum > pmf_Obs) && (inum > inum_min)
      pmf_table_inum = pmf_table_inum * (inum * (col1mRowSum + fnum + inum))
      inum = inum - 1.0
      pmf_table_inum = pmf_table_inum / ((rowsummfnum - inum) * (colsum(2) - inum))
   end
   while (pmf_table_inum <= pmf_Obs) && (inum <= mode)
      pmf_table_inum = pmf_table_inum * (rowsummfnum - inum) * (colsum(2) - inum)
      inum = inum + 1.0
      pmf_table_inum = pmf_table_inum / (inum * (col1mRowSum + fnum + inum))
   end
   pmf_table_jnum = pmf_hypergeometric(jnum, rowsummfnum, colsum(2), colsum1_2)
   while pmf_table_jnum = 0.0
      jnum = Int((jnum + mode) * 0.5)
      pmf_table_jnum = pmf_hypergeometric(jnum, rowsummfnum, colsum(2), colsum1_2)
   end
   while (pmf_table_jnum > pmf_Obs) && (jnum < jnum_max)
      pmf_table_jnum = pmf_table_jnum * ((rowsummfnum - jnum) * (colsum(2) - jnum))
      jnum = jnum + 1.0
      pmf_table_jnum = pmf_table_jnum / (jnum * (col1mRowSum + fnum + jnum))
   end
   while (pmf_table_jnum < pmf_Obs) && (jnum >= mode)
      pmf_table_jnum = pmf_table_jnum * (jnum * (col1mRowSum + fnum + jnum))
      jnum = jnum - 1.0
      pmf_table_jnum = pmf_table_jnum / ((rowsummfnum - jnum) * (colsum(2) - jnum))
   end
#Debug.Print fnum, d4num, dnum(5), dnum(6), inum, jnum, pmf_Obs
   inumstart = inum
   jnumstart = jnum
   if inum > jnum
      fet_26 = 1.0
      Exit Function
   end
   cdf = cdf_hypergeometric(inum - 1, rowsummfnum, colsum(2), colsum1_2)
   ccdf = comp_cdf_hypergeometric(jnum, rowsummfnum, colsum(2), colsum1_2)
   fnum_start = fnum
   continue_up = true
   continue_down = true
   inum_save = inum
   jnum_save = jnum
   pmf_table_inum_save = pmf_table_inum
   pmf_table_jnum_save = pmf_table_jnum
   cdf_save = cdf
   ccdf_save = ccdf
   cdf_start = cdf
   ccdf_start = ccdf
   prob = 0.0
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
         c_up(i) = true
         c_down(i) = true
      Next i
      Do
         probf = pmfd4 * (cdf + ccdf)
         pmf_Obs = pmf_Obs * pmfd4 # was pmf_Obs = pmf_Obs / (pmfd4 * pmf_prod(5))
         pmfd4_down = pmfd4
         pmfd4_up = pmfd4
         d4num_down = d4num
         d4num_up = d4num
         pmfd4_down = pmfd4_down * d4num_down * (colsum(3) - fnum + dnumsum(4) + d4num_down)
         d4num_down = d4num_down - 1
         pmfd4_down = pmfd4_down / ((fnum - dnumsum(4) - d4num_down) * (colsum(4) - d4num_down))
         pmfd4_up = pmfd4_up * ((fnum - dnumsum(4) - d4num_up) * (colsum(4) - d4num_up))
         d4num_up = d4num_up + 1.0
         pmfd4_up = pmfd4_up / (d4num_up * (colsum(3) - fnum + dnumsum(4) + d4num_up))
         inum = inum
         Do
            pmfd4 = max(pmfd4_down, pmfd4_up)
            if pmfd4 = 0.0 Exit Do
            if pmfd4 * pmf_table_inum <= pmf_Obs
               Do
                  cdf = cdf + pmf_table_inum
                  pmf_table_inum = pmf_table_inum * (rowsummfnum - inum) * (colsum(2) - inum)
                  inum = inum + 1.0
                  pmf_table_inum = pmf_table_inum / (inum * (col1mRowSum + fnum + inum))
               end Until (pmfd4 * pmf_table_inum > pmf_Obs) || (inum > mode)
            end
            if pmfd4 * pmf_table_jnum < pmf_Obs
               Do
                  ccdf = ccdf + pmf_table_jnum
                  pmf_table_jnum = pmf_table_jnum * (jnum * (col1mRowSum + fnum + jnum))
                  jnum = jnum - 1.0
                  pmf_table_jnum = pmf_table_jnum / ((rowsummfnum - jnum) * (colsum(2) - jnum))
               end Until (pmfd4 * pmf_table_jnum > pmf_Obs) || (jnum < mode)
            end
            if inum > jnum
               pmfd4_down = cdf_hypergeometric(d4num_down, fnum - dnumsum(4), colsum(4), colsumsum(4))
               pmfd4_up = comp_cdf_hypergeometric(d4num_up - 1.0, fnum - dnumsum(4), colsum(4), colsumsum(4))
               probf = probf + pmfd4_down + pmfd4_up
               Exit Do
            end
            probf = probf + pmfd4 * (cdf + ccdf)
            if pmfd4_down > pmfd4_up
               pmfd4_down = pmfd4_down * d4num_down * (colsum(3) - fnum + dnumsum(4) + d4num_down)
               d4num_down = d4num_down - 1.0
               pmfd4_down = pmfd4_down / ((fnum - dnumsum(4) - d4num_down) * (colsum(4) - d4num_down))
            else
               pmfd4_up = pmfd4_up * ((fnum - dnumsum(4) - d4num_up) * (colsum(4) - d4num_up))
               d4num_up = d4num_up + 1.0
               pmfd4_up = pmfd4_up / (d4num_up * (colsum(3) - fnum + dnumsum(4) + d4num_up))
            end
#Debug.Print inum, jnum, pmf_table_inum, pmf_table_jnum, cdf, ccdf, probf
         end
         tc = 5
         probf5(tc) = probf5(tc) + probf * pmfd(tc)
         Do
            Do
               el5 = true
               if c_up(tc)
                  pmfd(tc) = pmfd(tc) * ((fnum - dnumsum(tc - 1)) * (colsum(tc) - dnum(tc)))
                  dnum(tc) = dnum(tc) + 1.0
                  dnumsum(tc - 1) = dnumsum(tc) + dnum(tc)
                  pmfd(tc) = pmfd(tc) / (dnum(tc) * (colsumsum(tc - 1) - fnum + dnumsum(tc - 1)))
                  c_up(tc) = pmfd(tc) > 0.0
                  if Not c_up(tc)
                     pmfd(tc) = pmfd_save(tc)
                     dnum(tc) = dnum_save(tc)
                     dnumsum(tc - 1) = dnumsum(tc) + dnum(tc)
                     inum_next(tc) = inum_save5(tc)
                     jnum_next(tc) = jnum_save5(tc)
                     pmf_table_inum_next(tc) = pmf_table_inum_save5(tc)
                     pmf_table_jnum_next(tc) = pmf_table_jnum_save5(tc)
                     cdf_next(tc) = cdf_save5(tc)
                     ccdf_next(tc) = ccdf_save5(tc)
                  end
               end
               if Not c_up(tc)
                  pmfd(tc) = pmfd(tc) * dnum(tc) * (colsumsum(tc - 1) - fnum + dnumsum(tc - 1))
                  dnum(tc) = dnum(tc) - 1.0
                  dnumsum(tc - 1) = dnumsum(tc) + dnum(tc)
                  pmfd(tc) = pmfd(tc) / ((fnum - dnumsum(tc - 1)) * (colsum(tc) - dnum(tc)))
                  c_down(tc) = pmfd(tc) > 0.0
                  if Not c_down(tc)
                     tc = tc + 1
                     if tc > c Exit Do
                     probf5(tc) = probf5(tc) + probf5(tc - 1) * pmfd(tc)
#Debug.Print tc, probf5(tc), dnum(tc), probf5(tc - 1), pmfd(tc)
                     probf5(tc - 1) = 0.0
                     el5 = false
                  end
               end
            end Until el5
            if tc > c Exit Do
            
            row3sum = fnum - dnumsum(tc - 1)
            ml(3) = row3sum
            For i = 4 To tc - 1
               ml(i) = Int(row3sum * colsum(i) / colsumsum(tc - 1) + 0.5)
               ml(3) = ml(3) - ml(i)
            Next i
            
            Do #Update guess at mode vector
               all_d_zero = true
               For i = 3 To tc - 2
                  For j = i + 1 To tc - 1
                     d = ml(i) - Int((colsum(i) + 1.0) * (ml(i) + ml(j) + 1.0) / (colsum(i) + colsum(j) + 2.0))
                     if d <> 0.0
                        ml(i) = ml(i) - d
                        ml(j) = ml(j) + d
                        all_d_zero = false
                     end
                  Next j
               Next i
            end Until all_d_zero
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
#Debug.Print fnum, d4num, dnum(5), dnum(6), inum, jnum, pmf_Obs, pmff, pmfd4, pmfd(5), pmfd(6)
            if pmf_table_inum * (inum * (col1mRowSum + fnum + inum)) / ((rowsummfnum - inum + 1.0) * (colsum(2) - inum + 1.0)) > pmf_Obs
               fet_26 = "Problem with cdf"
            end
            while (pmf_table_inum <= pmf_Obs) && (inum <= mode)
               cdf = cdf + pmf_table_inum
               pmf_table_inum = pmf_table_inum * (rowsummfnum - inum) * (colsum(2) - inum)
               inum = inum + 1.0
               pmf_table_inum = pmf_table_inum / (inum * (col1mRowSum + fnum + inum))
            end
            if pmf_table_jnum * ((rowsummfnum - jnum) * (colsum(2) - jnum)) / ((jnum + 1.0) * (col1mRowSum + fnum + jnum + 1.0)) > pmf_Obs
               fet_26 = "Problem with ccdf"
            end
            while (pmf_table_jnum < pmf_Obs) && (jnum >= mode)
               ccdf = ccdf + pmf_table_jnum
               pmf_table_jnum = pmf_table_jnum * (jnum * (col1mRowSum + fnum + jnum))
               jnum = jnum - 1.0
               pmf_table_jnum = pmf_table_jnum / ((rowsummfnum - jnum) * (colsum(2) - jnum))
            end
            #el5 = true
            if inum > jnum
               if c_up(tc)
                  probf5(tc) = probf5(tc) + comp_cdf_hypergeometric(dnum(tc) - 1.0, fnum - dnumsum(tc), colsum(tc), colsumsum(tc))
                  pmfd(tc) = pmfd_save(tc)
                  dnum(tc) = dnum_save(tc)
                  dnumsum(tc - 1) = dnumsum(tc) + dnum(tc)
                  inum_next(tc) = inum_save5(tc)
                  jnum_next(tc) = jnum_save5(tc)
                  pmf_table_inum_next(tc) = pmf_table_inum_save5(tc)
                  pmf_table_jnum_next(tc) = pmf_table_jnum_save5(tc)
                  cdf_next(tc) = cdf_save5(tc)
                  ccdf_next(tc) = ccdf_save5(tc)
                  c_up(tc) = false
               else
                  probf5(tc) = probf5(tc) + cdf_hypergeometric(dnum(tc), fnum - dnumsum(tc), colsum(tc), colsumsum(tc))
                  c_down(tc) = false
                  tc = tc + 1
                  if tc > c Exit Do
                  probf5(tc) = probf5(tc) + probf5(tc - 1) * pmfd(tc)
#Debug.Print tc, probf5(tc), inum, jnum, dnum(tc), probf5(tc - 1), pmfd(tc)
                  probf5(tc - 1) = 0.0
               end
               el5 = false
            end
         end Until el5
         if tc > c Exit Do
         For i = tc To 5 Step -1
            inum_next(i) = inum
            jnum_next(i) = jnum
            pmf_table_inum_next(i) = pmf_table_inum
            pmf_table_jnum_next(i) = pmf_table_jnum
            cdf_next(i) = cdf
            ccdf_next(i) = ccdf
         Next i
         For i = tc - 1 To 5 Step -1
            c_up(i) = true
            c_down(i) = true
            inum_save5(i) = inum
            jnum_save5(i) = jnum
            pmf_table_inum_save5(i) = pmf_table_inum
            pmf_table_jnum_save5(i) = pmf_table_jnum
            cdf_save5(i) = cdf
            ccdf_save5(i) = ccdf
         Next i
      end

      prob = prob + probf5(c) * pmff
#Debug.Print c + 1, prob, fnum, probf5(c), pmff
      probf5(c) = 0.0
      if continue_up
         pmff = pmff * ((rowsummfnum) * (colsumsum(c) - fnum))
         fnum = fnum + 1.0
         rowsummfnum = rowsum - fnum
         pmff = pmff / ((colsum1_2 - rowsummfnum) * (fnum))
         continue_up = pmff > 0.0
         if Not continue_up
            pmff = pmff_start
            fnum = fnum_start
            rowsummfnum = rowsum - fnum
            inum_save = inumstart
            jnum_save = jnumstart
            cdf_save = cdf_start
            ccdf_save = ccdf_start
         else
            inum = inum_save
            jnum = jnum_save
            cdf = cdf_save
            ccdf = ccdf_save
            inum = inum - 1.0
            temp = PBB(col1mRowSum + fnum + inum, rowsummfnum - inum, colsum(2) - inum, inum + 1.0)
            hTerm = hTerm * ((rowsummfnum - inum) * (colsum(2) - inum) * (colsum1_2 + 1.0))
            inum = inum + 1.0
            hTerm = hTerm / ((rowsummfnum + 1.0) * (colsum(2) + 1.0) * (col1mRowSum + fnum + inum))
            pmf_table_inum = hTerm #pmf_hypergeometric(inum, rowsummfnum, colsum(2), colsum1_2)
            cdf = cdf + temp
            temp = PBB(jnum, colsum(2) - jnum, rowsummfnum - jnum + 1.0, col1mRowSum + fnum + jnum)
            hTerm = hTerm * (jnum * (col1mRowSum + fnum + jnum) * (colsum1_2 + 1.0))
            jnum = jnum - 1.0
            hTerm = hTerm / ((rowsummfnum + 1.0) * (colsum(2) - jnum) * (colsum(1) + 1.0))
            pmf_table_jnum = hTerm #pmf_hypergeometric(jnum, rowsummfnum, colsum(2), colsum1_2)
            ccdf = ccdf + temp
         end
      end
      Do
         if Not continue_up
            pmff = pmff * ((colsum1_2 - rowsummfnum) * (fnum))
            fnum = fnum - 1.0
            rowsummfnum = rowsum - fnum
            pmff = pmff / ((rowsummfnum) * (colsumsum(c) - fnum))
            continue_down = pmff > 0.0
            if Not continue_down Exit Do
            inum = inum_save
            jnum = jnum_save
            cdf = cdf_save
            ccdf = ccdf_save
            temp = PBB(inum, colsum(2) - inum, rowsummfnum - inum, col1mRowSum + fnum + inum + 1.0)
            hTerm = hTerm * (rowsummfnum - inum) * (colsum(2) - inum) * (colsum1_2 + 1.0)
            inum = inum + 1.0
            hTerm = hTerm / (inum * (colsum(1) + 1.0) * (colsum1_2 - rowsummfnum + 1.0))
            pmf_table_inum = hTerm #pmf_hypergeometric(inum, rowsummfnum, colsum(2), colsum1_2)
            cdf = cdf + temp
            temp = PBB(col1mRowSum + fnum + jnum + 1.0, rowsummfnum - jnum - 1.0, colsum(2) - jnum, jnum + 1.0)
            hTerm = hTerm * ((jnum + 1.0) * (col1mRowSum + fnum + jnum + 1.0) * (colsum1_2 + 1.0)) / ((colsum(2) + 1.0) * (rowsummfnum - jnum) * (colsum1_2 - rowsummfnum + 1.0))
            pmf_table_jnum = hTerm #pmf_hypergeometric(jnum, rowsummfnum, colsum(2), colsum1_2)
            ccdf = ccdf + temp
         end
         
         For i = c To 5 Step -1
            c_up(i) = true
            c_down(i) = true
            ml(i) = Int(fnum * colsum(i) / colsumsum(c) + 0.5)
            dnumsum(i - 1) = dnumsum(i) + ml(i)
         Next i
         ml(4) = Int(fnum * colsum(4) / colsumsum(c) + 0.5)
         ml(3) = fnum - ml(4) - dnumsum(4)
       
         Do #Update guess at mode vector
            all_d_zero = true
            For i = 3 To c - 1
               For j = i + 1 To c
                  d = ml(i) - Int((colsum(i) + 1.0) * (ml(i) + ml(j) + 1.0) / (colsum(i) + colsum(j) + 2.0))
                  if d <> 0.0
                     ml(i) = ml(i) - d
                     ml(j) = ml(j) + d
                     all_d_zero = false
                  end
               Next j
            Next i
         end Until all_d_zero
#Debug.Print ml(3), ml(4), ml(5), ml(6)
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
         mode = Int((rowsummfnum + 1.0) * (colsum(2) + 1.0) / (colsum1_2 + 2.0))
         #if ml(4) <> Int((fnum - dnumsum(4) + 1.0) * (colsum(4) + 1.0) / (colsumsum(4) + 2.0))
         #   fet_26 = "Problem with mode"
         #end
         pmfd4 = pmf_hypergeometric(d4num, fnum - dnumsum(4), colsum(4), colsumsum(4))
         pmf_Obs = pmf_Obs_save / (pmfd4 * pmf_prod(5))
         inum_min = max(0.0, -(fnum + col1mRowSum))
         jnum_max = min(rowsummfnum, colsum(2))
         if jnum_max > 0.0 && pmf_Obs < 1.0
            while (pmf_table_inum > pmf_Obs) && (inum > inum_min) || (inum > mode)
               pmf_table_inum = pmf_table_inum * (inum * (col1mRowSum + fnum + inum))
               inum = inum - 1.0
               pmf_table_inum = pmf_table_inum / ((rowsummfnum - inum) * (colsum(2) - inum))
               cdf = cdf - pmf_table_inum
            end
            while (pmf_table_inum < pmf_Obs) && (inum <= mode)
               cdf = cdf + pmf_table_inum
               pmf_table_inum = pmf_table_inum * (rowsummfnum - inum) * (colsum(2) - inum)
               inum = inum + 1.0
               pmf_table_inum = pmf_table_inum / (inum * (col1mRowSum + fnum + inum))
            end
            if inum = inum_min cdf = 0.0
            while (pmf_table_jnum > pmf_Obs) && (jnum < jnum_max) || (jnum < mode)
               pmf_table_jnum = pmf_table_jnum * ((rowsummfnum - jnum) * (colsum(2) - jnum))
               jnum = jnum + 1.0
               pmf_table_jnum = pmf_table_jnum / (jnum * (col1mRowSum + fnum + jnum))
               ccdf = ccdf - pmf_table_jnum
            end
            while (pmf_table_jnum < pmf_Obs) && (jnum >= mode)
               ccdf = ccdf + pmf_table_jnum
               pmf_table_jnum = pmf_table_jnum * (jnum * (col1mRowSum + fnum + jnum))
               jnum = jnum - 1.0
               pmf_table_jnum = pmf_table_jnum / ((rowsummfnum - jnum) * (colsum(2) - jnum))
            end
            if jnum = jnum_max ccdf = 0.0
         else
            inum = jnum + 1.0
         end
         exit_loop = true
         if inum > jnum
            if continue_up
               prob = prob + comp_cdf_hypergeometric(fnum - 1.0, colsumsum(c), rowsum, cs)
#Debug.Print c + 1, prob, fnum, comp_cdf_hypergeometric(fnum - 1.0, colsumsum(c), rowsum, cs)
               continue_up = false
               pmff = pmff_start
               fnum = fnum_start
               rowsummfnum = rowsum - fnum
               inum_save = inumstart
               jnum_save = jnumstart
               cdf_save = cdf_start
               ccdf_save = ccdf_start
               exit_loop = false
            else
               prob = prob + cdf_hypergeometric(fnum, colsumsum(c), rowsum, cs)
#Debug.Print c + 1, prob, fnum, cdf_hypergeometric(fnum, colsumsum(c), rowsum, cs)
               continue_down = false
            end
         end
      end Until exit_loop
      if Not continue_down Exit Do
#Debug.Print fnum, d4num, dnum(5), dnum(6), inum, jnum, pmf_Obs, pmff, pmfd4, pmfd(5), pmfd(6)
      cdf_save = cdf
      ccdf_save = ccdf
      inum_save = inum
      jnum_save = jnum
      pmf_table_inum_save = pmf_table_inum
      pmf_table_jnum_save = pmf_table_jnum
   end Until !continue_down
   fet_26 = prob
   Exit Function
elseif c >= 6
   Dim knum::Float64, prob_d::Float64
   knum = ml(c)
   pmff = pmf_hypergeometric(knum, colsum(c), rowsum, cs)
   pmff_start = pmff
   inum = inumstart
   jnum = jnumstart
   prob_d = fet_26(c - 1, colsum, rowsum - knum, pmf_Obs / pmff, inum, jnum)
   inumstart = inum
   jnumstart = jnum
   if inum > jnum
      fet_26 = 1.0
      Exit Function
   end
   prob = pmff * prob_d

   Do
      pmff = pmff * ((colsum(c) - knum) * (rowsum - knum))
      knum = knum + 1.0
      pmff = pmff / (knum * (cs - colsum(c) - rowsum + knum))
      if pmff = 0.0 Exit Do
      prob_d = fet_26(c - 1, colsum, rowsum - knum, pmf_Obs / pmff, inum, jnum)
      if inum > jnum
         prob = prob + comp_cdf_hypergeometric(knum - 1.0, colsum(c), rowsum, cs)
         Exit Do
      end
      prob = prob + pmff * prob_d
   end
   pmff = pmff_start
   knum = ml(c)
   inum = inumstart
   jnum = jnumstart
   Do
      pmff = pmff * (knum * (cs - colsum(c) - rowsum + knum))
      knum = knum - 1.0
      pmff = pmff / ((colsum(c) - knum) * (rowsum - knum))
      if pmff = 0.0 Exit Do
      prob_d = fet_26(c - 1, colsum, rowsum - knum, pmf_Obs / pmff, inum, jnum)
      if inum > jnum
         prob = prob + cdf_hypergeometric(knum, colsum(c), rowsum, cs)
         Exit Do
      end
      prob = prob + pmff * prob_d
   end
   fet_26 = prob
else
   fet_26 = "c must be >= 5"
end
end

function fet(r As Range)::Float64
Dim cs::Float64, rowsum::Float64, rs::Float64, maxc::Float64, d::Float64, a::Float64, b::Float64, c::Float64, bPlusc::Float64, cp::Float64, rsp::Float64, pmf_Obs::Float64, pmf_d::Float64, pmf_e::Float64
Dim pm::Float64, cd::Float64, prob::Float64, pmfh::Float64, pmfh_save::Float64, temp::Float64, inum::Float64, jnum::Float64, knum::Float64, knum_save::Float64
Dim rc As Long, cc As Long, i As Long, j As Long, k As Long
Dim inum_save::Float64, jnum_save::Float64, es11::Float64, es12::Float64, es12_save::Float64, es12p13::Float64, es01::Float64, es02::Float64, es023::Float64, pmf_d_save::Float64, pmf_e_save::Float64, prob_d::Float64, mode::Float64
Dim all_d_zero::Bool
rc = r.Rows.count
cc = r.Columns.count
if rc < 2 || cc < 2 || min(rc, cc) >= 3 && max(rc, cc) >= 4
   fet = [#VALUE!]
   Exit Function
end
#Change data so that it is 2x3 or 2x4 rather than 3x2 or 4x2.
if rc > cc
   i = rc
   rc = cc
   cc = i
   ReDim os(1 To rc, 1 To cc)::Float64, es(0 To rc, 0 To cc)::Float64, colsum(cc)::Float64
   For i = 1 To rc
       For j = 1 To cc
           os(i, j) = r.Item(j, i)
       Next j
   Next i
elseif cc * rc = 4 && false
   fet = old_fet_22(r.Item(1, 1), r.Item(1, 2), r.Item(2, 1), r.Item(2, 2))
   Exit Function
else
   ReDim os(1 To rc, 1 To cc)::Float64, es(0 To rc, 0 To cc)::Float64, colsum(cc)::Float64
   For i = 1 To rc
       For j = 1 To cc
           os(i, j) = r.Item(i, j)
       Next j
   Next i
end
#Calculate row totals and check that all values are non-negative integers
cs = 0.0
For i = 1 To rc
   rs = 0.0
   For j = 1 To cc
      if os(i, j) < 0 || Int(os(i, j)) <> os(i, j)
         fet = [#VALUE!]
         Exit Function
      end
      rs = rs + os(i, j)
   Next j
   if rs = 0.0
      fet = [#VALUE!]
      Exit Function
   end
   es(i, 0) = rs
   cs = cs + rs
Next i
es(0, 0) = cs
#Calculate column totals and find column with largest total
maxc = 0.0
For i = 1 To cc
   rs = 0.0
   For j = 1 To rc
      rs = rs + os(j, i)
   Next j
   es(0, i) = rs
   if maxc < rs
      k = i
      maxc = rs
   end
Next i
#Swap largest column into column 1
if k <> 1
   rs = es(0, 1)
   es(0, 1) = maxc
   es(0, k) = rs
   For j = 1 To rc
       rs = os(j, 1)
       os(j, 1) = os(j, k)
       os(j, k) = rs
   Next j
end
For i = 2 To cc - 1
   maxc = 0.0
   For j = i To cc
      if maxc < es(0, j)
         k = j
         maxc = es(0, j)
      end
   Next j
   if k <> i
      rs = es(0, i)
      es(0, i) = maxc
      es(0, k) = rs
      For j = 1 To rc
          rs = os(j, i)
          os(j, i) = os(j, k)
          os(j, k) = rs
      Next j
   end
Next i
if es(0, cc) = 0
   fet = [#VALUE!]
   Exit Function
end

if rc = 2
   if es(1, 0) < es(2, 0)
      For j = 1 To cc
          rs = os(1, j)
          os(1, j) = os(2, j)
          os(2, j) = rs
      Next j
      rs = es(1, 0)
      es(1, 0) = es(2, 0)
      es(2, 0) = rs
   end
   For j = 0 To cc
      colsum(j) = es(0, j)
   Next j
   rowsum = es(2, 0)
   #Guess at start points for most likely table
   rs = 0.0
   For i = 1 To rc
      For j = 1 To cc
         rsp = es(i, 0) * es(0, j) / cs
         rs = rs + (os(i, j) - rsp) ^ 2 / rsp
      Next j
   Next i
   knum = rowsum * colsum(2) / colsum(0)
   if cc > 2
      rsp = cs * (1.0 / rowsum) * (1.0 / colsum(cc - 1) + 1.0 / colsum(cc))
      pm = abs2(rs / rsp)
   else
      pm = abs(os(2, 2) - knum)
   end
   inum = Int(knum - pm + 0.5)
   jnum = Int(knum + pm + 0.5)
   pmf_Obs = 1.0
   For i = cc To 2 Step -1
      pmf_Obs = pmf_Obs * pmf_hypergeometric(os(2, i), colsum(i), rowsum, cs)
      rowsum = rowsum - os(2, i)
      cs = cs - colsum(i)
   Next i
   if pmf_Obs = 0.0
      fet = 0.0
      Exit Function
   end
   pmf_Obs = pmf_Obs * (1 + 0.000000000000001 * (10 - log(pmf_Obs)))
   if cc = 2
      fet = fet_22(cc, colsum, es(2, 0), pmf_Obs, inum, jnum)
   elseif cc = 3
      fet = fet_23(cc, colsum, es(2, 0), pmf_Obs, inum, jnum)
   elseif cc = 4
      fet = fet_24(cc, colsum, es(2, 0), pmf_Obs, inum, jnum)
   elseif cc >= 5
      fet = fet_25(cc, colsum, es(2, 0), pmf_Obs, inum, jnum)
   else
      fet = fet_26(cc, colsum, es(2, 0), pmf_Obs, inum, jnum)
   end
   Exit Function
elseif rc = 3
   if es(1, 0) > es(2, 0)
      For j = 1 To cc
          rs = os(1, j)
          os(1, j) = os(2, j)
          os(2, j) = rs
      Next j
      rs = es(1, 0)
      es(1, 0) = es(2, 0)
      es(2, 0) = rs
   end
   if es(2, 0) > es(3, 0)
      For j = 1 To cc
         rs = os(2, j)
         os(2, j) = os(3, j)
         os(3, j) = rs
      Next j
      rs = es(2, 0)
      es(2, 0) = es(3, 0)
      es(3, 0) = rs
      if es(1, 0) > es(2, 0)
         For j = 1 To cc
            rs = os(1, j)
            os(1, j) = os(2, j)
            os(2, j) = rs
         Next j
         rs = es(1, 0)
         es(1, 0) = es(2, 0)
         es(2, 0) = rs
      end
   end
#if sum of first row > sum of third column then transpose data
   if es(1, 0) > es(0, 3)
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
   end
elseif cc > 3
   fet = [#VALUE!]
   Exit Function
end

#Initial guess at mode
For i = 1 To cc
   es(rc, i) = es(0, i)
Next i
rs = 0.0
For i = 1 To rc
   if i < rc es(i, cc) = es(i, 0)
   For j = 1 To cc
      rsp = es(i, 0) * es(0, j) / cs
      rs = rs + (os(i, j) - rsp) ^ 2 / rsp
      if (i < rc) && (j < cc)
         es(i, j) = Int(rsp + 0.5)
         es(i, cc) = es(i, cc) - es(i, j)
         es(rc, j) = es(rc, j) - es(i, j)
      end
   Next j
   if i < rc es(rc, cc) = es(rc, cc) - es(i, cc)
Next i

es01 = es(0, 1)
es02 = es(0, 2)
es023 = es02 + es(0, 3)
pmf_d = pmf_hypergeometric(os(1, 1), es01, es(1, 0), cs)
pmf_e = pmf_hypergeometric(os(1, 2), es02, os(1, 2) + os(1, 3), es023)
pmf_Obs = pmf_d * pmf_e * pmf_hypergeometric(os(2, 1), os(2, 1) + os(3, 1), es(2, 0), es(2, 0) + es(3, 0)) * pmf_hypergeometric(os(2, 2), os(2, 2) + os(3, 2), os(2, 2) + os(2, 3), os(2, 2) + os(3, 2) + os(2, 3) + os(3, 3))
pmf_Obs = pmf_Obs * (1.0000000000001)
if pmf_Obs = 0.0
   fet = 0.0
   Exit Function
end


#Refining guess for mode
Do
   all_d_zero = true
   d = es(2, 2) - Int((es(2, 2) + es(2, 3) + 1.0) * (es(2, 2) + es(3, 2) + 1.0) / (es(2, 2) + es(2, 3) + es(3, 2) + es(3, 3) + 2.0))
   if d <> 0
      es(2, 2) = es(2, 2) - d
      es(2, 3) = es(2, 3) + d
      es(3, 2) = es(3, 2) + d
      es(3, 3) = es(3, 3) - d
      all_d_zero = false
   end
   d = es(1, 2) - Int((es(1, 2) + es(1, 3) + 1.0) * (es(1, 2) + es(3, 2) + 1.0) / (es(1, 2) + es(1, 3) + es(3, 2) + es(3, 3) + 2.0))
   if d <> 0
      es(1, 2) = es(1, 2) - d
      es(1, 3) = es(1, 3) + d
      es(3, 2) = es(3, 2) + d
      es(3, 3) = es(3, 3) - d
      all_d_zero = false
   end
   d = es(1, 2) - Int((es(1, 2) + es(1, 3) + 1.0) * (es(1, 2) + es(2, 2) + 1.0) / (es(1, 2) + es(1, 3) + es(2, 2) + es(2, 3) + 2.0))
   if d <> 0
      es(1, 2) = es(1, 2) - d
      es(1, 3) = es(1, 3) + d
      es(2, 2) = es(2, 2) + d
      es(2, 3) = es(2, 3) - d
      all_d_zero = false
   end
   d = es(2, 1) - Int((es(2, 1) + es(2, 3) + 1.0) * (es(2, 1) + es(3, 1) + 1.0) / (es(2, 1) + es(2, 3) + es(3, 1) + es(3, 3) + 2.0))
   if d <> 0
      es(2, 1) = es(2, 1) - d
      es(2, 3) = es(2, 3) + d
      es(3, 1) = es(3, 1) + d
      es(3, 3) = es(3, 3) - d
      all_d_zero = false
   end
   d = es(1, 1) - Int((es(1, 1) + es(1, 3) + 1.0) * (es(1, 1) + es(3, 1) + 1.0) / (es(1, 1) + es(1, 3) + es(3, 1) + es(3, 3) + 2.0))
   if d <> 0
      es(1, 1) = es(1, 1) - d
      es(1, 3) = es(1, 3) + d
      es(3, 1) = es(3, 1) + d
      es(3, 3) = es(3, 3) - d
      all_d_zero = false
   end
   d = es(1, 1) - Int((es(1, 1) + es(1, 3) + 1.0) * (es(1, 1) + es(2, 1) + 1.0) / (es(1, 1) + es(1, 3) + es(2, 1) + es(2, 3) + 2.0))
   if d <> 0
      es(1, 1) = es(1, 1) - d
      es(1, 3) = es(1, 3) + d
      es(2, 1) = es(2, 1) + d
      es(2, 3) = es(2, 3) - d
      all_d_zero = false
   end
   d = es(2, 1) - Int((es(2, 1) + es(2, 2) + 1.0) * (es(2, 1) + es(3, 1) + 1.0) / (es(2, 1) + es(2, 2) + es(3, 1) + es(3, 2) + 2.0))
   if d <> 0
      es(2, 1) = es(2, 1) - d
      es(2, 2) = es(2, 2) + d
      es(3, 1) = es(3, 1) + d
      es(3, 2) = es(3, 2) - d
      all_d_zero = false
   end
   d = es(1, 1) - Int((es(1, 1) + es(1, 2) + 1.0) * (es(1, 1) + es(3, 1) + 1.0) / (es(1, 1) + es(1, 2) + es(3, 1) + es(3, 2) + 2.0))
   if d <> 0
      es(1, 1) = es(1, 1) - d
      es(1, 2) = es(1, 2) + d
      es(3, 1) = es(3, 1) + d
      es(3, 2) = es(3, 2) - d
      all_d_zero = false
   end
   d = es(1, 1) - Int((es(1, 1) + es(1, 2) + 1.0) * (es(1, 1) + es(2, 1) + 1.0) / (es(1, 1) + es(1, 2) + es(2, 1) + es(2, 2) + 2.0))
   if d <> 0
      es(1, 1) = es(1, 1) - d
      es(1, 2) = es(1, 2) + d
      es(2, 1) = es(2, 1) + d
      es(2, 2) = es(2, 2) - d
      all_d_zero = false
   end
end Until all_d_zero
#Guess at start points for most likely table
rsp = cs * (1.0 / es(2, 0) + 1.0 / es(3, 0)) * (1.0 / es(0, 2) + 1.0 / es(0, 3))
pm = abs2(rs / rsp)

knum = es(2, 0) * es(0, 2) / es(0, 0)
inum = Int(knum - pm + 0.5)
jnum = Int(knum + pm + 0.5)
inum_save = inum
jnum_save = jnum
#Work from mode out
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

prob = 0.0
prob_d = fet_23(3, colsum, rowsum, pmf_Obs / (pmf_d * pmf_e), inum, jnum)
if inum > jnum
   fet = 1.0
   Exit Function
end
prob_d = prob_d * pmf_e

while pmf_d > 0.0
   Do
      colsum(2) = colsum(2) - 1.0
      colsum(3) = colsum(3) + 1.0
      pmf_e = pmf_e * ((es12p13 - es12) * (es02 - es12))
      es12 = es12 + 1.0
      pmf_e = pmf_e / (es12 * (es023 - es02 - es12p13 + es12))
      if pmf_e = 0.0 Exit Do
      prob_d = prob_d + pmf_e * fet_23(3, colsum, rowsum, pmf_Obs / (pmf_d * pmf_e), inum, jnum)
      if inum > jnum
         prob_d = prob_d + comp_cdf_hypergeometric(es12, es02, es12p13, es023)
         Exit Do
      end
   end
   inum = inum_save
   jnum = jnum_save
   es12 = es12_save
   colsum(2) = es02 - es12
   colsum(3) = colsum(0) - colsum(1) - colsum(2)
   pmf_e = pmf_e_save
   Do
      colsum(2) = colsum(2) + 1.0
      colsum(3) = colsum(3) - 1.0
      pmf_e = pmf_e * (es12 * (es023 - es02 - es12p13 + es12))
      es12 = es12 - 1.0
      pmf_e = pmf_e / ((es12p13 - es12) * (es02 - es12))
      if pmf_e = 0.0 Exit Do
      prob_d = prob_d + pmf_e * fet_23(3, colsum, rowsum, pmf_Obs / (pmf_d * pmf_e), inum, jnum)
      if inum > jnum
         prob_d = prob_d + cdf_hypergeometric(es12 - 1.0, es02, es12p13, es023)
         Exit Do
      end
   end
   prob = prob + prob_d * pmf_d
   inum = inum_save
   jnum = jnum_save
   pmf_d = pmf_d * ((es01 - es11) * (es(1, 0) - es11))
   es11 = es11 + 1.0
   es12p13 = es12p13 - 1.0
   pmf_d = pmf_d / (es11 * (cs - es01 - es(1, 0) + es11))
   if pmf_d = 0.0 Exit Do
   es12 = Int((es02 + 1.0) * (es12p13 + 1.0) / (es023 + 2.0))
   colsum(1) = es01 - es11
   colsum(2) = es02 - es12
   colsum(3) = colsum(0) - colsum(1) - colsum(2)
#Guess of mode values for second row
   #Es[1,1] = es11 Don#t touch Es[1,1]
   es(1, 2) = es12
   es(1, 3) = es(1, 0) - es11 - es12
   es(2, 1) = Int(rowsum * colsum(1) / colsum(0) + 0.5)
   es(2, 2) = Int(rowsum * colsum(2) / colsum(0) + 0.5)
   es(2, 3) = rowsum - es(2, 1) - es(2, 2)
   es(3, 1) = colsum(1) - es(2, 1)
   es(3, 2) = colsum(2) - es(2, 2)
   es(3, 3) = colsum(3) - es(2, 3)
   
#Refining guess for mode with Es[1,1] fixed.
   Do
      all_d_zero = true
      d = es(2, 2) - Int((es(2, 2) + es(2, 3) + 1.0) * (es(2, 2) + es(3, 2) + 1.0) / (es(2, 2) + es(2, 3) + es(3, 2) + es(3, 3) + 2.0))
      if d <> 0
         es(2, 2) = es(2, 2) - d
         es(2, 3) = es(2, 3) + d
         es(3, 2) = es(3, 2) + d
         es(3, 3) = es(3, 3) - d
         all_d_zero = false
      end
      d = es(1, 2) - Int((es(1, 2) + es(1, 3) + 1.0) * (es(1, 2) + es(3, 2) + 1.0) / (es(1, 2) + es(1, 3) + es(3, 2) + es(3, 3) + 2.0))
      if d <> 0
         es(1, 2) = es(1, 2) - d
         es(1, 3) = es(1, 3) + d
         es(3, 2) = es(3, 2) + d
         es(3, 3) = es(3, 3) - d
         all_d_zero = false
      end
      d = es(1, 2) - Int((es(1, 2) + es(1, 3) + 1.0) * (es(1, 2) + es(2, 2) + 1.0) / (es(1, 2) + es(1, 3) + es(2, 2) + es(2, 3) + 2.0))
      if d <> 0
         es(1, 2) = es(1, 2) - d
         es(1, 3) = es(1, 3) + d
         es(2, 2) = es(2, 2) + d
         es(2, 3) = es(2, 3) - d
         all_d_zero = false
      end
      d = es(2, 1) - Int((es(2, 1) + es(2, 3) + 1.0) * (es(2, 1) + es(3, 1) + 1.0) / (es(2, 1) + es(2, 3) + es(3, 1) + es(3, 3) + 2.0))
      if d <> 0
         es(2, 1) = es(2, 1) - d
         es(2, 3) = es(2, 3) + d
         es(3, 1) = es(3, 1) + d
         es(3, 3) = es(3, 3) - d
         all_d_zero = false
      end
      d = es(2, 1) - Int((es(2, 1) + es(2, 2) + 1.0) * (es(2, 1) + es(3, 1) + 1.0) / (es(2, 1) + es(2, 2) + es(3, 1) + es(3, 2) + 2.0))
      if d <> 0
         es(2, 1) = es(2, 1) - d
         es(2, 2) = es(2, 2) + d
         es(3, 1) = es(3, 1) + d
         es(3, 2) = es(3, 2) - d
         all_d_zero = false
      end
   end Until all_d_zero
   
   es12 = es(1, 2)
   es12_save = es12
   colsum(2) = es02 - es12
   colsum(3) = colsum(0) - colsum(1) - colsum(2)
   pmf_e = pmf_hypergeometric(es12, es02, es12p13, es023)
   pmf_e_save = pmf_e
   prob_d = fet_23(3, colsum, rowsum, pmf_Obs / (pmf_d * pmf_e), inum, jnum)
   if inum > jnum
      prob = prob + comp_cdf_hypergeometric(es11 - 1.0, es01, es(1, 0), cs)
      Exit Do
   end
   prob_d = prob_d * pmf_e
   inum_save = inum
   jnum_save = jnum
end
inum = Int(knum - pm + 0.5)
jnum = Int(knum + pm + 0.5)
pmf_d = pmf_d_save
es11 = es(1, 1)
es12 = es(1, 2)
es12p13 = es(1, 0) - es11
Do
   pmf_d = pmf_d * (es11 * (cs - es01 - es(1, 0) + es11))
   es11 = es11 - 1.0
   es12p13 = es12p13 + 1.0
   pmf_d = pmf_d / ((es01 - es11) * (es(1, 0) - es11))
   if pmf_d = 0.0 Exit Do
   es12 = Int((es02 + 1.0) * (es12p13 + 1.0) / (es023 + 2.0))
   colsum(1) = es01 - es11
   colsum(2) = es02 - es12
   colsum(3) = colsum(0) - colsum(1) - colsum(2)
#Guess of mode values for second row
   #Es[1,1] = es11 Don#t touch Es[1,1]
   es(1, 2) = es12
   es(1, 3) = es(1, 0) - es11 - es12
   es(2, 1) = Int(rowsum * colsum(1) / colsum(0) + 0.5)
   es(2, 2) = Int(rowsum * colsum(2) / colsum(0) + 0.5)
   es(2, 3) = rowsum - es(2, 1) - es(2, 2)
   es(3, 1) = colsum(1) - es(2, 1)
   es(3, 2) = colsum(2) - es(2, 2)
   es(3, 3) = colsum(3) - es(2, 3)
   
#Refining guess for mode with Es[1,1] fixed.
   Do
      all_d_zero = true
      d = es(2, 2) - Int((es(2, 2) + es(2, 3) + 1.0) * (es(2, 2) + es(3, 2) + 1.0) / (es(2, 2) + es(2, 3) + es(3, 2) + es(3, 3) + 2.0))
      if d <> 0
         es(2, 2) = es(2, 2) - d
         es(2, 3) = es(2, 3) + d
         es(3, 2) = es(3, 2) + d
         es(3, 3) = es(3, 3) - d
         all_d_zero = false
      end
      d = es(1, 2) - Int((es(1, 2) + es(1, 3) + 1.0) * (es(1, 2) + es(3, 2) + 1.0) / (es(1, 2) + es(1, 3) + es(3, 2) + es(3, 3) + 2.0))
      if d <> 0
         es(1, 2) = es(1, 2) - d
         es(1, 3) = es(1, 3) + d
         es(3, 2) = es(3, 2) + d
         es(3, 3) = es(3, 3) - d
         all_d_zero = false
      end
      d = es(1, 2) - Int((es(1, 2) + es(1, 3) + 1.0) * (es(1, 2) + es(2, 2) + 1.0) / (es(1, 2) + es(1, 3) + es(2, 2) + es(2, 3) + 2.0))
      if d <> 0
         es(1, 2) = es(1, 2) - d
         es(1, 3) = es(1, 3) + d
         es(2, 2) = es(2, 2) + d
         es(2, 3) = es(2, 3) - d
         all_d_zero = false
      end
      d = es(2, 1) - Int((es(2, 1) + es(2, 3) + 1.0) * (es(2, 1) + es(3, 1) + 1.0) / (es(2, 1) + es(2, 3) + es(3, 1) + es(3, 3) + 2.0))
      if d <> 0
         es(2, 1) = es(2, 1) - d
         es(2, 3) = es(2, 3) + d
         es(3, 1) = es(3, 1) + d
         es(3, 3) = es(3, 3) - d
         all_d_zero = false
      end
      d = es(2, 1) - Int((es(2, 1) + es(2, 2) + 1.0) * (es(2, 1) + es(3, 1) + 1.0) / (es(2, 1) + es(2, 2) + es(3, 1) + es(3, 2) + 2.0))
      if d <> 0
         es(2, 1) = es(2, 1) - d
         es(2, 2) = es(2, 2) + d
         es(3, 1) = es(3, 1) + d
         es(3, 2) = es(3, 2) - d
         all_d_zero = false
      end
   end Until all_d_zero
   
   es12 = es(1, 2)
   es12_save = es12
   colsum(2) = es02 - es12
   colsum(3) = colsum(0) - colsum(1) - colsum(2)
   pmf_e = pmf_hypergeometric(es12, es02, es12p13, es023)
   pmf_e_save = pmf_e
   prob_d = fet_23(3, colsum, rowsum, pmf_Obs / (pmf_d * pmf_e), inum, jnum)
   if inum > jnum
      prob = prob + cdf_hypergeometric(es11, es01, es(1, 0), cs)
      Exit Do
   end
   prob_d = prob_d * pmf_e
   inum_save = inum
   jnum_save = jnum
   
   Do
      colsum(2) = colsum(2) - 1.0
      colsum(3) = colsum(3) + 1.0
      pmf_e = pmf_e * ((es12p13 - es12) * (es02 - es12))
      es12 = es12 + 1.0
      pmf_e = pmf_e / (es12 * (es023 - es02 - es12p13 + es12))
      if pmf_e = 0.0 Exit Do
      prob_d = prob_d + pmf_e * fet_23(3, colsum, rowsum, pmf_Obs / (pmf_d * pmf_e), inum, jnum)
      if inum > jnum
         prob_d = prob_d + comp_cdf_hypergeometric(es12, es02, es12p13, es023)
         Exit Do
      end
   end
   inum = inum_save
   jnum = jnum_save
   es12 = es12_save
   pmf_e = pmf_e_save
   colsum(2) = es02 - es12
   colsum(3) = colsum(0) - colsum(1) - colsum(2)
   Do
      colsum(2) = colsum(2) + 1.0
      colsum(3) = colsum(3) - 1.0
      pmf_e = pmf_e * (es12 * (es023 - es02 - es12p13 + es12))
      es12 = es12 - 1.0
      pmf_e = pmf_e / ((es12p13 - es12) * (es02 - es12))
      if pmf_e = 0.0 Exit Do
      prob_d = prob_d + pmf_e * fet_23(3, colsum, rowsum, pmf_Obs / (pmf_d * pmf_e), inum, jnum)
      if inum > jnum
         prob_d = prob_d + cdf_hypergeometric(es12 - 1.0, es02, es12p13, es023)
         Exit Do
      end
   end
   prob = prob + prob_d * pmf_d
   inum = inum_save
   jnum = jnum_save
end

fet = prob
end

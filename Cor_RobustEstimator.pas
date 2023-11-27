unit Cor_RobustEstimator;
// Delphi adaptation by Bruce Eglington (2006/07/10)
// of the C code for TRobustEstimator.cxx,v 1.12 2006/05/16,
// written by Anna Kreshuk  08/10/2004
//
//*****************************************************************************
// * Original version Copyright (C) 1995-2004, Rene Brun and Fons Rademakers. *
// * All rights reserved.                                                     *
// *                                                                          *
// ****************************************************************************
//
//////////////////////////////////////////////////////////////////////////////
//
//  TRobustEstimator
//
// Minimum Covariance Determinant Estimator - a Fast Algorithm
// invented by Peter J.Rousseeuw and Katrien Van Dreissen
// "A Fast Algorithm for the Minimum covariance Determinant Estimator"
// Technometrics, August 1999, Vol.41, NO.3
//
// What are robust estimators?
// "An important property of an estimator is its robustness. An estimator
// is called robust if it is insensitive to measurements that deviate
// from the expected behaviour. There are 2 ways to treat such deviating
// measurements: one may either try to recongize them and then remove
// them from the data sample; or one may leave them in the sample, taking
// care that they do not influence the estimate unduly. In both cases robust
// estimators are needed...Robust procedures compensate for systematic errors
// as much as possible, and indicate any situation in which a danger of not being
// able to operate reliably is detected."
// R.Fruhwirth, M.Regler, R.K.Bock, H.Grote, D.Notz
// "Data Analysis Techniques for High-Energy Physics", 2nd edition
//
// What does this algorithm do?
// It computes a highly robust estimator of multivariate location and scatter.
// Then, it takes those estimates to compute robust distances of all the
// data vectors. Those with large robust distances are considered outliers.
// Robust distances can then be plotted for better visualization of the data.
//
// How does this algorithm do it?
// The MCD objective is to find h observations(out of n) whose classical
// covariance matrix has the lowest determinant. The MCD estimator of location
// is then the average of those h points and the MCD estimate of scatter
// is their covariance matrix. The minimum(and default) h = (n+nvariables+1)/2
// so the algorithm is effective when less than (n+nvar+1)/2 variables are outliers.
// The algorithm also allows for exact fit situations - that is, when h or more
// observations lie on a hyperplane. Then the algorithm still yields the MCD location T
// and scatter matrix S, the latter being singular as it should be. From (T,S) the
// program then computes the equation of the hyperplane.
//
// How can this algorithm be used?
// In any case, when contamination of data is suspected, that might influence
// the classical estimates.
// Also, robust estimation of location and scatter is a tool to robustify
// other multivariate techniques such as, for example, principal-component analysis
// and discriminant analysis.
//
//
//
//
// Technical details of the algorithm:
// 0.The default h = (n+nvariables+1)/2, but the user may choose any interger h with
//   (n+nvariables+1)/2<=h<=n. The program then reports the MCD's breakdown value
//   (n-h+1)/n. If you are sure that the dataset contains less than 25% contamination
//   which is usually the case, a good compromise between breakdown value and
//  efficiency is obtained by putting h=[.75*n].
// 1.If h=n,the MCD location estimate is the average of the whole dataset, and
//   the MCD scatter estimate is its covariance matrix. Report this and stop
// 2.If nvariables=1 (univariate data), compute the MCD estimate by the exact
//   algorithm of Rousseeuw and Leroy (1987, pp.171-172) in O(nlogn)time and stop
// 3.From here on, h<n and nvariables>=2.
//   3a.If n is small:
//    - repeat (say) 500 times:
//    -- construct an initial h-subset, starting from a random (nvar+1)-subset
//    -- carry out 2 C-steps (described in the comments of CStep funtion)
//    - for the 10 results with lowest det(S):
//    -- carry out C-steps until convergence
//    - report the solution (T, S) with the lowest det(S)
//   3b.If n is larger (say, n>600), then
//    - construct up to 5 disjoint random subsets of size nsub (say, nsub=300)
//    - inside each subset repeat 500/5 times:
//    -- construct an initial subset of size hsub=[nsub*h/n]
//    -- carry out 2 C-steps
//    -- keep the best 10 results (Tsub, Ssub)
//    - pool the subsets, yielding the merged set (say, of size nmerged=1500)
//    - in the merged set, repeat for each of the 50 solutions (Tsub, Ssub)
//    -- carry out 2 C-steps
//    -- keep the 10 best results
//    - in the full dataset, repeat for those best results:
//    -- take several C-steps, using n and h
//    -- report the best final result (T, S)
// 4.To obtain consistency when the data comes from a multivariate normal
//   distribution, covariance matrix is multiplied by a correction factor
// 5.Robust distances for all elements, using the final (T, S) are calculated
//   Then the very final mean and covariance estimates are calculated only for
//   values, whose robust distances are less than a cutoff value (0.975 quantile
//   of chi2 distribution with nvariables degrees of freedom)
//
//////////////////////////////////////////////////////////////////////////////

interface

uses Math, NumRecipes_varb;

const
  tNMax = 20;
  tMMax = 5;
  nbest = 10;
  nmini = 2; {need to check and change this *****************************}
  kChiMedian : array[1..50] of double = (
         0.454937, 1.38629, 2.36597, 3.35670, 4.35146, 5.34812, 6.34581, 7.34412, 8.34283,
         9.34182, 10.34, 11.34, 12.34, 13.34, 14.34, 15.34, 16.34, 17.34, 18.34, 19.34,
        20.34, 21.34, 22.34, 23.34, 24.34, 25.34, 26.34, 27.34, 28.34, 29.34, 30.34,
        31.34, 32.34, 33.34, 34.34, 35.34, 36.34, 37.34, 38.34, 39.34, 40.34,
        41.34, 42.34, 43.34, 44.34, 45.34, 46.34, 47.34, 48.34, 49.33);
  kChiQuant : array [1..50] of double = (
         5.02389, 7.3776,9.34840,11.1433,12.8325,
        14.4494,16.0128,17.5346,19.0228,20.4831,21.920,23.337,
        24.736,26.119,27.488,28.845,30.191,31.526,32.852,34.170,
        35.479,36.781,38.076,39.364,40.646,41.923,43.194,44.461,
        45.722,46.979,48.232,49.481,50.725,51.966,53.203,54.437,
        55.668,56.896,58.120,59.342,60.561,61.777,62.990,64.201,
        65.410,66.617,67.821,69.022,70.222,71.420);

type
  bbb = double;
  tfData = array[1..tNMax, 1..tMMax] of double;
  Tsscp = array[0..tMMax+1, 0..tMMax+1] of double;
  TvecdoubleN = array[1..tNMax+1] of double;
  TvecintegerN = array[1..tNMax+1] of integer;
  TvecdoubleM = array[1..tMMax+1] of double;
  TvecintegerM = array[1..tMMax+1] of integer;
  Tvecdoublenbest = array[1..nbest] of double;
  Tmstock = array[1..nbest,1..tMMax] of double;
  Tcstock = array[1..tMMax,1..tMMax*nbest] of double;
  Tindex = array[1..tNMax] of integer;

var
  maxind : integer;
  fRd : TvecdoubleN;
  fData : TfData;
  fCovariance : Tsscp;

procedure EvaluateRobustEstimator(data : Tfdata; fH, fN, fnvar : integer;
                              var Covariance : Tsscp;
                              var Correlation : Tsscp;
                              var Mean : TvecdoubleM);
procedure KOrdStat(ntotal : integer; a : TvecdoubleM;
                   k : integer; work : Tindex);
procedure ClearSscp(var sscp : Tsscp; M : integer);
procedure AddToSscp(M : integer; var sscp : Tsscp; temp : TvecdoubleM);
procedure CoVar  (     X          : TfData;
                   var A          : Tsscp;
                       N, M       : integer);
procedure CoVar2(sscp : Tsscp;
                N, M : integer;
            var Mean : TvecdoubleM;
            var Cov : Tsscp);
procedure Determinant(fCovariance : Tsscp;
                  var det : double;
                      N, fnvar : integer);
procedure Classic(data : Tfdata; N, M : integer;
              var Covariance : Tsscp;
              var Correlation : Tsscp;
              var Mean : TvecdoubleM);
procedure  Correl(Covariance : Tsscp;
              var Correlation : Tsscp;
                  M : integer);
procedure EvaluateUni(M : integer;
                      data : Tfdata;
                  var mean : double;
                  var sigma : double;
                      hh : integer);
function GetBDPoint(N, H : integer): integer;
function GetChiQuant(i : integer): double;
procedure CreateSubset(N, H, p : integer;
                       IIndex : TIndex;
                   var data : Tfdata;
                   var sscp : Tsscp;
                       ndist : TvecdoubleM);
procedure CStep (ntotal, htotal : integer; iindex : Tindex;
                 fnvar : integer;
                 data : Tfdata; sscp : Tsscp; fMean : TvecdoubleM;
             var ndist : TvecdoubleM;
             var det : double);
procedure CreateOrtSubset(var dat : Tfdata;
                              IIndex : TIndex;
                              hmerged : integer;
                              nmerged : integer;
                          var sscp : Tsscp;
                              ndist : TvecdoubleM);
procedure Exact(var ndist : TvecdoubleM);
procedure Exact2(var mstockbig : Tsscp;
                 var cstockbig : Tsscp;
                 var hyperplane : Tsscp;
                 deti : double;
                 nbest : integer;
                 kgroup : integer;
                 var sscp : Tsscp;
                 ndist : TvecdoubleM);
procedure Partition(nmini : integer; indsubdat : TvecintegerM);
procedure RDist(var sscp : Tsscp);
procedure RDraw(subdat : TvecdoubleM; ngroup : integer; indsubdat : TvecIntegerM);
function LocMax(nbest : integer; deti : Tvecdoublenbest) : integer;

implementation

procedure EvaluateRobustEstimator(data : Tfdata; fH, fN, fnvar : integer;
                              var Covariance : Tsscp;
                              var Correlation : Tsscp;
                              var Mean : TvecdoubleM);
const
  kEps = 1.0e-14;
  nmini = 300;
  k1=500;
var
  i, j, k : integer;
  ii, jj : integer;
  index : integer;
  iindex : TIndex;
  sscp : Tsscp;
  vec : TvecdoubleM;
  tmpindex : TIndex;
  tmpndist : TvecdoubleN;
  det : double;
  deti : Tvecdoublenbest;
  mstock : Tmstock;
  cstock : Tcstock;
  fExact : double;
  ndist : TvecdoubleM;
  fMean : TvecdoubleM;
begin
//Finds the estimate of multivariate mean and variance
   if (fH = fN) then
   begin
//     Warning("Evaluate","Chosen h = #observations, so classic estimates of location and scatter will be calculated");
     Classic(data, fN, fNvar, Covariance, Correlation, Mean);
     Halt;
   end;
   for i := 0 to nbest do
   begin
     deti[i] := 1.0e16;
   end;
   for i := 0 to fN do
   begin
     fRd[i] := 0.0;
   end;
   {for small n}
   if (fN<nmini*2) then
   begin
     {for storing the best fMeans and covariances}
     for k := 0 to k1 do
     begin
       CreateSubset(fN, fH, fNvar, tmpindex, fData, sscp, ndist);
       //calculate the mean and covariance of the created subset
       ClearSscp(sscp,fnvar);
       for i := 0 to fH do
       begin
         for j := 0 to fnvar do
         begin
           vec[j] := fData[tmpindex[i],j];
           AddToSscp(fnvar, sscp, vec);
         end;
         Covar2(sscp, fN, fnvar, fMean, fCovariance);
         Determinant(fCovariance,det,fN,fnvar);
         if (det < kEps) then
         begin
           Exact(ndist);
           Exit;
         end;
         //make 2 CSteps
         CStep(fN, fH, iindex, fnvar, fData, sscp, fMean, ndist, det);
         if (det < kEps) then
         begin
           Exact(ndist);
           Halt;
         end;
         Determinant(sscp,det,fN,fnvar);
         {
         det := CStep(fN, fH, index, fData, sscp, ndist);
         }
         if (det < kEps) then
         begin
           Exact(ndist);
           Halt;
         end else
         begin
            maxind := LocMax(nbest,deti);
            if(det < deti[maxind]) then
            begin
               deti[maxind] := det;
               for ii := 0 to fNvar do
               begin
                 mstock[maxind,ii] := fMean[ii];
                 for jj := 0 to fNvar do
                 begin
                   cstock[ii,jj+maxind*fNvar] := fCovariance[ii,jj];
                 end;
               end;
            end;
         end;
       end;
     end;
   end;
      //now for nbest best results perform CSteps until convergence
      for i := 0 to nbest do
      begin
        for ii := 0 to fNvar do
        begin
          fMean[ii] := mstock[i,ii];
          for jj := 0 to fNvar do
          begin
            Covariance[ii,jj] := cstock[ii,jj+i*fNvar];
          end;
        end;
      end;

(*

         det=1;
         while (det>kEps) {
            det=CStep(fN, fH, index, fData, sscp, ndist);
            if(TMath::Abs(det-deti[i])<kEps)
               break;
            else
               deti[i]=det;
         }
         for(ii=0; ii<fNvar; ii++) {
            mstock(i,ii)=fMean(ii);
            for (jj=0; jj<fNvar; jj++)
               cstock(ii,jj+i*fNvar)=fCovariance(ii, jj);
         }
      }

      Int_t detind=TMath::LocMin(nbest, deti);
      for(ii=0; ii<fNvar; ii++) {
         fMean(ii)=mstock(detind,ii);

         for(jj=0; jj<fNvar; jj++)
            fCovariance(ii, jj)=cstock(ii,jj+detind*fNvar);
      }

      if (deti[detind]!=0) {
         //calculate robust distances and throw out the bad points
         Int_t nout = RDist(sscp);
         Double_t cutoff=kChiQuant[fNvar-1];

         fOut.Set(nout);

         j=0;
         for (i=0; i<fN; i++) {
            if(fRd(i)>cutoff) {
               fOut[j]=i;
               j++;
            }
         }

      } else {
         fExact=Exact(ndist);
      }
      delete [] index;
      delete [] ndist;
      delete [] deti;
      return;

   }
   /////////////////////////////////////////////////
  //if n>nmini, the dataset should be partitioned
  //partitioning
  ////////////////////////////////////////////////
   Int_t indsubdat[5];
   Int_t nsub;
   for (ii=0; ii<5; ii++)
      indsubdat[ii]=0;

   nsub = Partition(nmini, indsubdat);

   Int_t sum=0;
   for (ii=0; ii<5; ii++)
      sum+=indsubdat[ii];
   Int_t *subdat=new Int_t[sum];
   RDraw(subdat, nsub, indsubdat);
   //now the indexes of selected cases are in the array subdat
   //matrices to store best means and covariances
   Int_t nbestsub=nbest*nsub;
   TMatrixD mstockbig(nbestsub, fNvar);
   TMatrixD cstockbig(fNvar, fNvar*nbestsub);
   TMatrixD hyperplane(nbestsub, fNvar);
   for (i=0; i<nbestsub; i++) {
      for(j=0; j<fNvar; j++)
         hyperplane(i,j)=0;
   }
   Double_t *detibig = new Double_t[nbestsub];
   Int_t maxind;
   maxind=TMath::LocMax(5, indsubdat);
   TMatrixD dattemp(indsubdat[maxind], fNvar);

   Int_t k2=Int_t(k1/nsub);
   //construct h-subsets and perform 2 CSteps in subgroups

   for (Int_t kgroup=0; kgroup<nsub; kgroup++) {
      //printf("group #%d\n", kgroup);
      Int_t ntemp=indsubdat[kgroup];
      Int_t temp=0;
      for (i=0; i<kgroup; i++)
         temp+=indsubdat[i];
      Int_t par;

      for(i=0; i<ntemp; i++) {
         for (j=0; j<fNvar; j++) {
            dattemp(i,j)=fData[subdat[temp+i]][j];
         }
      }
      Int_t htemp=Int_t(fH*ntemp/fN);

      for (i=0; i<nbest; i++)
         deti[i]=1e16;

      for(k=0; k<k2; k++) {
         CreateSubset(ntemp, htemp, fNvar, index, dattemp, sscp, ndist);
         ClearSscp(sscp);
         for (i=0; i<htemp; i++) {
            for(j=0; j<fNvar; j++) {
               vec(j)=dattemp(index[i],j);
            }
            AddToSscp(sscp, vec);
         }
         Covar(sscp, fMean, fCovariance, fSd, htemp);
         det = fCovariance.Determinant();
         if (det<kEps) {
            par =Exact2(mstockbig, cstockbig, hyperplane, deti, nbest, kgroup, sscp,ndist);
            if(par==nbest+1) {

               delete [] detibig;
               delete [] deti;
               delete [] subdat;
               delete [] ndist;
               delete [] index;
               return;
            } else
               deti[par]=det;
         } else {
            det = CStep(ntemp, htemp, index, dattemp, sscp, ndist);
            if (det<kEps) {
               par=Exact2(mstockbig, cstockbig, hyperplane, deti, nbest, kgroup, sscp, ndist);
               if(par==nbest+1) {

                  delete [] detibig;
                  delete [] deti;
                  delete [] subdat;
                  delete [] ndist;
                  delete [] index;
                  return;
               } else
                  deti[par]=det;
            } else {
               det=CStep(ntemp,htemp, index, dattemp, sscp, ndist);
               if(det<kEps){
                  par=Exact2(mstockbig, cstockbig, hyperplane, deti, nbest, kgroup, sscp,ndist);
                  if(par==nbest+1) {

                     delete [] detibig;
                     delete [] deti;
                     delete [] subdat;
                     delete [] ndist;
                     delete [] index;
                     return;
                  } else {
                     deti[par]=det;
                  }
               } else {
                  maxind=TMath::LocMax(nbest, deti);
                  if(det<deti[maxind]) {
                     deti[maxind]=det;
                     for(i=0; i<fNvar; i++) {
                        mstockbig(nbest*kgroup+maxind,i)=fMean(i);
                        for(j=0; j<fNvar; j++) {
                           cstockbig(i,nbest*kgroup*fNvar+maxind*fNvar+j)=fCovariance(i,j);

                        }
                     }
                  }

               }
            }
         }

         maxind=TMath::LocMax(nbest, deti);
         if (deti[maxind]<kEps)
            break;
      }


      for(i=0; i<nbest; i++) {
         detibig[kgroup*nbest + i]=deti[i];

      }

   }

   //now the arrays mstockbig and cstockbig store nbest*nsub best means and covariances
   //detibig stores nbest*nsub their determinants
   //merge the subsets and carry out 2 CSteps on the merged set for all 50 best solutions

   TMatrixD datmerged(sum, fNvar);
   for(i=0; i<sum; i++) {
      for (j=0; j<fNvar; j++)
         datmerged(i,j)=fData[subdat[i]][j];
   }
   //  printf("performing calculations for merged set\n");
   Int_t hmerged=Int_t(sum*fH/fN);

   Int_t nh;
   for(k=0; k<nbestsub; k++) {
      //for all best solutions perform 2 CSteps and then choose the very best
      for(ii=0; ii<fNvar; ii++) {
         fMean(ii)=mstockbig(k,ii);
         for(jj=0; jj<fNvar; jj++)
            fCovariance(ii, jj)=cstockbig(ii,k*fNvar+jj);
      }
      if(detibig[k]==0) {
         for(i=0; i<fNvar; i++)
            fHyperplane(i)=hyperplane(k,i);
         CreateOrtSubset(datmerged,index, hmerged, sum, sscp, ndist);

      }
      det=CStep(sum, hmerged, index, datmerged, sscp, ndist);
      if (det<kEps) {
         nh= Exact(ndist);
         if (nh>=fH) {
            fExact = nh;

            delete [] detibig;
            delete [] deti;
            delete [] subdat;
            delete [] ndist;
            delete [] index;
            return;
         } else {
            CreateOrtSubset(datmerged, index, hmerged, sum, sscp, ndist);
         }
      }

      det=CStep(sum, hmerged, index, datmerged, sscp, ndist);
      if (det<kEps) {
         nh=Exact(ndist);
         if (nh>=fH) {
            fExact = nh;
            delete [] detibig;
            delete [] deti;
            delete [] subdat;
            delete [] ndist;
            delete [] index;
            return;
         }
      }
      detibig[k]=det;
      for(i=0; i<fNvar; i++) {
         mstockbig(k,i)=fMean(i);
         for(j=0; j<fNvar; j++) {
            cstockbig(i,k*fNvar+j)=fCovariance(i, j);
         }
      }
   }
   //now for the subset with the smallest determinant
   //repeat CSteps until convergence
   Int_t minind=TMath::LocMin(nbestsub, detibig);
   det=detibig[minind];
   for(i=0; i<fNvar; i++) {
      fMean(i)=mstockbig(minind,i);
      fHyperplane(i)=hyperplane(minind,i);
      for(j=0; j<fNvar; j++)
         fCovariance(i, j)=cstockbig(i,minind*fNvar + j);
   }
   if(det<kEps)
      CreateOrtSubset(fData, index, fH, fN, sscp, ndist);
   det=1;
   while (det>kEps) {
      det=CStep(fN, fH, index, fData, sscp, ndist);
      if(TMath::Abs(det-detibig[minind])<kEps) {
         break;
      } else {
         detibig[minind]=det;
      }
   }
   if(det<kEps) {
      Exact(ndist);
      fExact=kTRUE;
   }
   Int_t nout = RDist(sscp);
   Double_t cutoff=kChiQuant[fNvar-1];

   fOut.Set(nout);

   j=0;
   for (i=0; i<fN; i++) {
      if(fRd(i)>cutoff) {
         fOut[j]=i;
         j++;
      }
   }

   delete [] detibig;
   delete [] deti;
   delete [] subdat;
   delete [] ndist;
   delete [] index;
   return;
}


*)


end;

procedure CStep (ntotal, htotal : integer; iindex : Tindex;
                 fnvar : integer;
                 data : Tfdata; sscp : Tsscp; fMean : TvecdoubleM;
             var ndist : TvecdoubleM;
             var det : double);
  //from the input htotal-subset constructs another htotal subset with lower determinant
  //
  //As proven by Peter J.Rousseeuw and Katrien Van Driessen, if distances for all elements
  //are calculated, using the formula:d_i=Sqrt((x_i-M)*S_inv*(x_i-M)), where M is the mean
  //of the input htotal-subset, and S_inv - the inverse of its covariance matrix, then
  //htotal elements with smallest distances will have covariance matrix with determinant
  //less or equal to the determinant of the input subset covariance matrix.
  //
  //determinant for this htotal-subset with smallest distances is returned
var
  i , j : integer;
  vec, temp : TvecdoubleM;
  fSD : Tsscp;
begin
  for j := 0 to ntotal do
  begin
    ndist[j] := 0.0;
    for i := 0 to fnvar do
    begin
      temp[i] := data[j,i] - fMean[i];
    end;
    {temp*=fInvcovariance;}
    for i := 0 to fnvar do
    begin
      ndist[j] := ndist[j] + (data[j,i] - fMean[i]) * temp[i];
    end;
  end;
  //taking h smallest
  KOrdStat(ntotal, ndist, htotal-1, iindex);
  //writing their mean and covariance
  ClearSscp(sscp,fnvar);
  for i := 0 to htotal do
  begin
    for j := 0 to fnvar do
    begin
      temp[j] := data[iindex[i],j];
    end;
    AddToSscp(fnvar, sscp, temp);
  end;
  {
  Covar(sscp, ntotal, fNvar, fMean, fCovariance, fSd);
  Determinant(fCovariance,det, ntotal, fnvar);
  }
end;

procedure KOrdStat(ntotal : integer; a : TvecdoubleM;
                   k : integer; work : Tindex);
const
  kWorkMax = 100;
type
  TworkLocal = array[1..kWorkMax] of integer;
var
  isAllocated : boolean;
  i, ir, j, l, mid, ii, rk : integer;
  arr : integer;
  ind : Tindex;
  temp : integer;
  tmp : double;
  workLocal : TworkLocal;
  Int_t : Tindex;
begin
  //because I need an Int_t work array
  if (work[k] > 0) then
  begin
    ind[k] := work[k];
  end else
  begin
    ind[k] := workLocal[k];
    if (ntotal > kWorkMax) then
    begin
      isAllocated := true;
    end;
  end;
  for ii := 0 to ntotal do
  begin
    ind[ii] := ii;
  end;
  rk := k;
  l := 0;
  ir := ntotal-1;
  repeat
    {for(;;)}
    if (ir <= l+1) then
    begin
      //active partition contains 1 or 2 elements
      if ((ir = l+1) and (a[ind[ir]] < a[ind[l]])) then
      begin
        temp := ind[l];
        ind[l] := ind[ir];
        ind[ir] := temp;
      end;
      tmp := a[ind[rk]];
    end else
    begin
      mid := l + ir;
      if (mid > 1) then //choose median of left, center and right
      begin
        temp := ind[mid];
        ind[mid] := ind[l+1];
        ind[l+1] := temp;//elements as partitioning element arr.
      end;
      if (a[ind[l]] > a[ind[ir]]) then
      begin
        //also rearrange so that a[l]<=a[l+1]
        temp := ind[l];
        ind[l] := ind[ir];
        ind[ir] := temp;
      end;
      if (a[ind[l+1]] > a[ind[ir]]) then
      begin
        temp := ind[l+1];
        ind[l+1] := ind[ir];
        ind[ir] := temp;
      end;
      if (a[ind[l]] > a[ind[l+1]]) then
      begin
        temp := ind[l];
        ind[l] := ind[l+1];
        ind[l+1] := temp;
      end;
      i := l+1;        //initialize pointers for partitioning
      j := ir;
      arr := ind[l+1];
      repeat
        {for (;;)}
        i := i +1;
        while (a[ind[i]] < a[arr]) do
        begin
          j := j -1;
          while (a[ind[j]] > a[arr]) do
          begin
            if (j < i) then
            begin
              break;  //pointers crossed, partitioning complete
            end else
            begin
              temp := ind[i];
              ind[i] := ind[j];
              ind[j] := temp;
            end;
          end;
        end;
      until (i > j); {not correct until}
      ind[l+1] := ind[j];
      ind[j] := arr;
      if (j>=rk) then ir := j-1; //keep active the partition that
      if (j<=rk) then l := i;    //contains the k_th element
    end;
  until (i > j); {not correct until}
end;

procedure ClearSscp(var sscp : Tsscp; M : integer);
var
  i, j : integer;
begin
  //clear the sscp matrix, used for covariance and mean calculation
  for i := 0 to M+1 do
  begin
    for j := 0 to M+1 do
    begin
      sscp[i,j] := 0.0;
    end;
  end;
end;

procedure AddToSscp(M : integer; var sscp : Tsscp; temp : TvecdoubleM);
var
  i, j : integer;
begin
  for j := 1 to M+1 do
  begin
    sscp[0,j] := sscp[0,j] + temp[j-1];
    sscp[j,0] := sscp[0,j];
  end;
  for i := 1 to M+1 do
  begin
    for j := 1 to M+1 do
    begin
      sscp[i,j] := sscp[i,j] + temp[i-1] * temp[j-1];
    end;
  end;
end;

procedure CoVar  (     X          : TfData;
                   var A          : Tsscp;
                       N, M       : integer);
var
  I, J, K     :  integer;
  temp, temp1,
  SX1, SX2,
  SX1X2       : double;
begin
  temp:=1.0*N;
  temp1:=temp-1.0;
  for I:=1 to M do
  begin
    for J:=1 to M do
    begin
      SX1:=0.0;
      SX2:=0.0;
      SX1X2:=0.0;
      for K:=1 to N do
      begin
        SX1:=SX1+X[K,I];
        SX2:=SX2+X[K,J];
        SX1X2:=SX1X2+X[K,I]*X[K,J];
      end;
      A[I,J]:=(SX1X2-SX1*SX2/temp)/temp1;
      A[J,I]:=A[I,J];
    end;
  end;
end;

procedure CoVar2(sscp : Tsscp;
                N, M : integer;
            var Mean : TvecdoubleM;
            var Cov : Tsscp);
var
  i, j : integer;
  f : double;
  tSd : TvecdoubleM;
begin
  //calculates mean and covariance
  for i := 0 to M do
  begin
    Mean[i] := sscp[0,i+1];
    tSd[i] := sscp[i+1,i+1];
    f := (tSd[i]-Mean[i]*Mean[i]/N)/(N-1);
    if (f > 1.0e-14) then tSd[i] := sqrt(f)
                     else tSd[i] := 0.0;
    Mean[i] := Mean[i]/N;
  end;
  for i := 0 to M do
  begin
    for j := 0 to M do
    begin
      cov[i,j] := sscp[i+1,j+1]-N*Mean[i]*Mean[j];
      cov[i,j] := cov[i,j]/(N-1);
    end;
  end;
end;

procedure Determinant(fCovariance : Tsscp;
                  var det : double;
                      N, fnvar : integer);
begin
  {}
end;

procedure Classic(data : Tfdata; N, M : integer;
              var Covariance : Tsscp;
              var Correlation : Tsscp;
              var Mean : TvecdoubleM);
  //called when h=n. Returns classic covariance matrix
  //and mean
var
  i, j : integer;
  temp : TvecdoubleM;
  sscp : Tsscp;
begin
  ClearSscp(sscp,M);
  for i := 0 to N do
  begin
    for j := 0 to M do
    begin
      temp[j] := Data[i,j];
    end;
    AddToSscp(M, sscp, temp);
  end;
  Covar(data, Covariance, N, M);
  Correl(Covariance, Correlation, M);
end;

procedure  Correl(Covariance : Tsscp;
              var Correlation : Tsscp;
                  M : integer);
var
  i, j : integer;
  sd : TvecdoubleM;
begin
  for j := 0 to M do
  begin
    sd[j] := 1.0/sqrt(Covariance[j,j]);
  end;
  for i := 0 to M do
  begin
    for j := 0 to M do
    begin
      if (i = j) then
      begin
        Correlation[i,j] := 1.0;
      end else
      begin
        Correlation[i,j] := Covariance[i,j] * sd[i]*sd[j];
      end;
    end;
  end;
end;

procedure EvaluateUni(M : integer;
                      data : Tfdata;
                  var mean : double;
                  var sigma : double;
                      hh : integer);
begin
  {}
end;

function GetBDPoint(N, H : integer): integer;
begin
  Result := (N - H + 1) div N;
end;

function GetChiQuant(i : integer): double;
begin
  if ((i < 0) or (i > 50)) then
  begin
    Result := 0.0;
  end else
  begin
    Result := kchiQuant[i];
  end;
end;

procedure CreateSubset(N, H, p : integer;
                       IIndex : TIndex;
                   var data : Tfdata;
                   var sscp : Tsscp;
                       ndist : TvecdoubleM);
begin
  {}
end;

procedure CreateOrtSubset(var dat : Tfdata;
                              IIndex : TIndex;
                              hmerged : integer;
                              nmerged : integer;
                          var sscp : Tsscp;
                              ndist : TvecdoubleM);
begin
  {}
end;

procedure Exact(var ndist : TvecdoubleM);
begin
  {}
end;

procedure Exact2(var mstockbig : Tsscp;
                 var cstockbig : Tsscp;
                 var hyperplane : Tsscp;
                 deti : double;
                 nbest : integer;
                 kgroup : integer;
                 var sscp : Tsscp;
                 ndist : TvecdoubleM);
begin
  {}
end;

procedure Partition(nmini : integer; indsubdat : TvecintegerM);
begin
  {}
end;

procedure RDist(var sscp : Tsscp);
begin
  {}
end;

procedure RDraw(subdat : TvecdoubleM; ngroup : integer; indsubdat : TvecIntegerM);
begin
  {}
end;

function LocMax(nbest : integer; deti : Tvecdoublenbest) : integer;
begin
  {}
end;



end.

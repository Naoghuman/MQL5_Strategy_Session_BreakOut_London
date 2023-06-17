/*
 * Copyright (C) 2023 Naoghuman
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this program.  If not, see <http://www.gnu.org/licenses/>.
 */
/* ============================================================================================
 * DESCRIPTION
 * This script export an overview about the performance from the strategy 'Session-BreakOut London'.
 * 
 * Analyse will be the performance from the first BreakOut (LONG/SHORT9 and if in the same day
 * available the second BreakOut.
 * 
 * ============================================================================================
 */
#property copyright "Naoghuman (Peter Rogge)"
#property link      "https://github.com/Naoghuman/MQL5_Strategy_Session_BreakOut_London/blob/main/MQL5/Scripts/_Naoghuman/_GitHub/Session-BreakOut_London/MT5_Script_Report_Session-Breakout-London.mq5"
#property version   "1.001"

// Input
// display window of the input parameters during the script's launch
#property script_show_inputs

input group "== 'Session-Breakout London' ======================="
input string InputMarket = "AUDCHF"; // List of markets (separated with ;) which will be analysed
//input string InputMarket = "AUDCAD;AUDCHF;AUDJPY;AUDNZD;AUDUSD;CADCHF;CADJPY;CHFJPY;EURAUD;EURCAD;EURCHF;EURGBP;EURJPY;EURNZD;EURUSD;GBPAUD;GBPCAD;GBPCHF;GBPJPY;GBPNZD;GBPUSD;NZDCAD;NZDCHF;NZDJPY;NZDUSD;USDCHF;USDJPY"; // List of markets (separated with ;) which will be analysed
input bool Trade_LONG  = true; // Trade LONG positions
input bool Trade_SHORT = true; // Trade SHORT positions

input group "== Time-Management ==============================="
input int             CandlesToAnalyse  = 150000;    // Candles to analyse in this report
input ENUM_TIMEFRAMES DelayAfterTrade   = PERIOD_M1; // Delay after a Trade
input int             BreakOutHour      = 9;         // BreakOut Hour (Local Time) from Session London
input int             BreakOutMinute    = 0;         // - BreakOut Minute (Local Time) from Session London
input int             BreakOutCandlesNr = 1;         // Candles (before the BreakOut) defines the BreakOut-Range

input group "== Report-Management ============================="
input bool ShowReportSingleTradeData = false; // Shows (on/off) the data from the single Trades

// Enums
enum ENUM_BREAKOUT
{
   NONE  = 0, // No Breakout
   LONG  = 1, // LONG Breakout
   SHORT = 2  // SHORT Breakout
};

// Includes

// Defines

// Enums

// Input variables

// Global variables
ENUM_BREAKOUT _1stBreakOutType;
double _1stEntry_Price;
double _1stStopLoss_Price;
double _1stStopLoss_R;
double _1stMaxPotential_Price;
double _1stMaxPotential_R;
double _1stMaxPotential_R_All;
double _1stLosser_Weekday_LONG[6];
double _1stLosser_Weekday_SHORT[6];
double _1stWinner_Weekday_LONG[6];
double _1stWinner_Weekday_SHORT[6];
double _1stResult_MPR_Weekday_LONG[6];
double _1stResult_MPR_Weekday_SHORT[6];

ENUM_BREAKOUT _2ndBreakOutType;
double _2ndEntry_Price;
double _2ndStopLoss_Price;
double _2ndStopLoss_R;
double _2ndMaxPotential_Price;
double _2ndMaxPotential_R;
double _2ndMaxPotential_R_All;
double _2ndLosser_Weekday_LONG[6];
double _2ndLosser_Weekday_SHORT[6];
double _2ndWinner_Weekday_LONG[6];
double _2ndWinner_Weekday_SHORT[6];
double _2ndResult_MPR_Weekday_LONG[6];
double _2ndResult_MPR_Weekday_SHORT[6];

void OnStart()
{
   // ## Scan every market ####################################################################
   // Scan every market
   string _Markets[];
   StringSplit(InputMarket, ';', _Markets);
   
   int _SizeMarkets = ArraySize(_Markets);
   for (int _Index = 0; _Index < _SizeMarkets; _Index++)
   {
      // ## Open file ############################################################################
      Comment("");
      
      string _Market            = _Markets[_Index];
      string _SpreadSheet       = "/Session-Breakout/Session-Breakout_London_" + _Market + "_" + TimeToString(TimeLocal(), TIME_DATE) + ".csv";
      int    _SpreadSheetHandle = FileOpen(_SpreadSheet, FILE_READ | FILE_WRITE | FILE_CSV | FILE_UNICODE | FILE_COMMON);
   
      FileSeek(_SpreadSheetHandle, 0, SEEK_END);
      
      // ## Write header #########################################################################
      if (ShowReportSingleTradeData)
      {
         FileWrite(_SpreadSheetHandle, "Market", "Date", "Weekday", "Time (Local)", "Time (Server)", "BreakOut (High)", "BreakOut (Low)",
                   "1. Position (Type)", "1. Entry (Price)", "1. Stop/Loss (Price)", "1. Stop/Loss (R)", "1. Max Potenial (Price)", "1. Max Potenial (R)",
                   "2. Position (Type)", "2. Entry (Price)", "2. Stop/Loss (Price)", "2. Stop/Loss (R)", "2. Max Potenial (Price)", "2. Max Potenial (R)");
      }
      
      // ## Data #################################################################################
      _1stMaxPotential_R_All = 0;
      _2ndMaxPotential_R_All = 0;
      
      ArrayFill(_1stLosser_Weekday_LONG,  0, 6, 0);
      ArrayFill(_1stLosser_Weekday_SHORT, 0, 6, 0);
      ArrayFill(_2ndLosser_Weekday_LONG,  0, 6, 0);
      ArrayFill(_2ndLosser_Weekday_SHORT, 0, 6, 0);
      
      ArrayFill(_1stWinner_Weekday_LONG,  0, 6, 0);
      ArrayFill(_1stWinner_Weekday_SHORT, 0, 6, 0);
      ArrayFill(_2ndWinner_Weekday_LONG,  0, 6, 0);
      ArrayFill(_2ndWinner_Weekday_SHORT, 0, 6, 0);
      
      ArrayFill(_1stResult_MPR_Weekday_LONG,  0, 6, 0);
      ArrayFill(_1stResult_MPR_Weekday_SHORT, 0, 6, 0);
      ArrayFill(_2ndResult_MPR_Weekday_LONG,  0, 6, 0);
      ArrayFill(_2ndResult_MPR_Weekday_SHORT, 0, 6, 0);
      
      // Check for every candle
      int _TradeCounter = 0;
      for (int i = CandlesToAnalyse; i > 64; i--)
      {
         MqlDateTime _MqlDateTime;
         TimeToStruct(iTime(_Market, PERIOD_M15, i), _MqlDateTime);
         
         // Check if we have a TradeDay (Monday - Friday).
         int _DayOfWeek;
         if (!IsTradeDay(_MqlDateTime, _DayOfWeek)) { continue; }
         
         // Check if we have the SessionBreakOut time
         if (!IsSessionBreakOutTime(_MqlDateTime)) { continue; }
      
         // Extract the BreakOut-Range
         double _BreakOutHigh_Price;
         double _BreakOutLow_Price;
         ExtractBreakOutRange(_Market, i, _BreakOutHigh_Price, _BreakOutLow_Price);
         
         // Check BreakOut
         int _BreakOutIndex;
         ENUM_BREAKOUT _FirstBreakOut = CheckFirstBreakOut(_Market, i, _BreakOutHigh_Price, _BreakOutLow_Price, _BreakOutIndex);
         if (_FirstBreakOut == NONE) { continue; }
         
         // Compute BreakOut data
         if (Trade_LONG && _FirstBreakOut == LONG)
         {
            ComputeBreakOutData_LONG(_Market, i, _DayOfWeek, _BreakOutHigh_Price, _BreakOutLow_Price, _BreakOutIndex);
            
            if (ShowReportSingleTradeData)
            {
               WriteTradeData(_SpreadSheetHandle, _Market, _DayOfWeek, _BreakOutHigh_Price, _BreakOutLow_Price, _BreakOutIndex);
            }
         }
         
         if (Trade_SHORT && _FirstBreakOut == SHORT)
         {
            ComputeBreakOutData_SHORT(_Market, i, _DayOfWeek, _BreakOutHigh_Price, _BreakOutLow_Price, _BreakOutIndex);
            
            if (ShowReportSingleTradeData)
            {
               WriteTradeData(_SpreadSheetHandle, _Market, _DayOfWeek, _BreakOutHigh_Price, _BreakOutLow_Price, _BreakOutIndex);
            }
         }
      }

      WriteTradeDataSummary(_SpreadSheetHandle, _Market);
      
      // ## Close file ###########################################################################
      FileClose(_SpreadSheetHandle);
   }
   
   Comment("DONE"); 
}

bool IsTradeDay(MqlDateTime &_MqlDateTime, int &_DayOfWeek)
{
   _DayOfWeek = _MqlDateTime.day_of_week;
   if (_DayOfWeek >= 1 && _DayOfWeek <= 5)
   {
      return true;
   }
   
   return false;
}

bool IsSessionBreakOutTime(MqlDateTime &_MqlDateTime)
{
   int _Hour = _MqlDateTime.hour;
   int _Min  = _MqlDateTime.min;
   if (_Hour == (BreakOutHour + 1) && _Min == BreakOutMinute)
   {
      return true;
   }
   
   return false;
}

void ExtractBreakOutRange(string _Market, int _Index, double &_BreakOutHigh_Price, double &_BreakOutLow_Price)
{
   /*
    * _Index = 09:00
    */
   _BreakOutHigh_Price = -7654321;
   _BreakOutLow_Price  =  7654321;
   for (int i = _Index + BreakOutCandlesNr; i > _Index; i--)
   {
      double _High        = iHigh(_Market, PERIOD_M15, i);
      _BreakOutHigh_Price = MathMax(_High, _BreakOutHigh_Price);
      
      double _Low        = iLow(_Market, PERIOD_M15, i);
      _BreakOutLow_Price = MathMin(_Low, _BreakOutLow_Price);
   }
}

ENUM_BREAKOUT CheckFirstBreakOut(string _Market, int _Index, double _BreakOutHigh_Price, double _BreakOutLow_Price, int &_BreakOutIndex)
{
   ENUM_BREAKOUT _BreakOut = NONE;
   for (int i = _Index; i > _Index - 48; i--)
   {
      // Check BreakOut (direction).
      double _iHigh = iHigh(_Market, PERIOD_M15, i);
      double _iLow  = iLow( _Market, PERIOD_M15, i);
      if (_iHigh > _BreakOutHigh_Price) // POSITION_TYPE_BUY
      {
         _BreakOut      = LONG;
         _BreakOutIndex = i;
   
         break;
      }
      
      if (_iLow  < _BreakOutLow_Price) // POSITION_TYPE_SELL
      {
         _BreakOut      = SHORT;
         _BreakOutIndex = i;
   
         break;
      }
   }
   
   return _BreakOut;
}

void ComputeBreakOutData_LONG(string _Market, int _Index, int _Weekday, double _BreakOutHigh_Price, double _BreakOutLow_Price, int _BreakOutIndex)
{
   // #########################################################################################
   _1stBreakOutType       = LONG;
   _1stEntry_Price        = -7654321;
   _1stStopLoss_Price     = -7654321;
   _1stStopLoss_R         =        0;
   _1stMaxPotential_Price = -7654321;
   _1stMaxPotential_R     =        0;
   
   _2ndBreakOutType       = NONE;
   _2ndEntry_Price        = 7654321;
   _2ndStopLoss_Price     = 7654321;
   _2ndStopLoss_R         =       0;
   _2ndMaxPotential_Price = 7654321;
   _2ndMaxPotential_R     =       0;
   
   // 1st BreakOut (POSITION_TYPE_BUY)
   _1stEntry_Price        = _BreakOutHigh_Price;
   _1stStopLoss_Price     = _BreakOutLow_Price;
   _1stStopLoss_R         = _BreakOutHigh_Price - _BreakOutLow_Price;
   
   bool _IsStoppedOut    = false;
   int  _StoppedOutIndex = 0;
   for (int i = _BreakOutIndex; i >= _Index - 48; i--)
   {    
      // Compute _MaxPotential_Price, _MaxPotenial_R
      double _High           = iHigh(  _Market, PERIOD_M15, i);
      _1stMaxPotential_Price = MathMax(_1stMaxPotential_Price, _High);
            
      // Check if we are stopped out
      double _Low = iLow(_Market, PERIOD_M15, i);
      if (_Low < _1stStopLoss_Price)
      {
         _IsStoppedOut    = true;
         _StoppedOutIndex = i;
         break;
      }
   }
       _1stMaxPotential_R  = (_1stMaxPotential_Price - _1stEntry_Price) / _1stStopLoss_R;
   if (_1stMaxPotential_R >= 1)
   {
      _1stWinner_Weekday_LONG[0]            += 1;
      _1stWinner_Weekday_LONG[_Weekday]     += 1;
      _1stResult_MPR_Weekday_LONG[0]        += _1stMaxPotential_R;
      _1stResult_MPR_Weekday_LONG[_Weekday] += _1stMaxPotential_R;
   }
   else
   {
      _1stLosser_Weekday_LONG[0]            += 1;
      _1stLosser_Weekday_LONG[_Weekday]     += 1;
      _1stResult_MPR_Weekday_LONG[0]        -= 1;
      _1stResult_MPR_Weekday_LONG[_Weekday] -= 1;
   }
   _1stMaxPotential_R_All += _1stMaxPotential_R;
   
   // #########################################################################################
   // 2nd BreakOut (POSITION_TYPE_SELL)
   if (!_IsStoppedOut) { return; }
   
   _2ndBreakOutType   = SHORT;
   _2ndEntry_Price    = _BreakOutLow_Price;
   _2ndStopLoss_Price = _BreakOutHigh_Price;
   _2ndStopLoss_R     = _BreakOutHigh_Price - _BreakOutLow_Price;
   
   for (int i = _StoppedOutIndex; i > _Index - 48; i--)
   {    
      // Compute _MaxPotential_Price, _MaxPotenial_R
      double _Low            = iLow(   _Market, PERIOD_M15, i);
      _2ndMaxPotential_Price = MathMin(_2ndMaxPotential_Price, _Low);
            
      // Check if we are stopped out
      double _High = iHigh(_Market, PERIOD_M15, i);
      if (_High > _2ndStopLoss_Price) { break; }
   }
       _2ndMaxPotential_R  = (_2ndEntry_Price - _2ndMaxPotential_Price) / _2ndStopLoss_R;
   if (_2ndMaxPotential_R >= 1)
   {
      _2ndWinner_Weekday_SHORT[0]            += 1;
      _2ndWinner_Weekday_SHORT[_Weekday]     += 1;
      _2ndResult_MPR_Weekday_SHORT[0]        += _1stMaxPotential_R;
      _2ndResult_MPR_Weekday_SHORT[_Weekday] += _1stMaxPotential_R;
   }
   else
   {
      _2ndLosser_Weekday_SHORT[0]            += 1;
      _2ndLosser_Weekday_SHORT[_Weekday]     += 1;
      _2ndResult_MPR_Weekday_SHORT[0]        -= 1;
      _2ndResult_MPR_Weekday_SHORT[_Weekday] -= 1;
   }
   _2ndMaxPotential_R_All += _2ndMaxPotential_R;
}

void ComputeBreakOutData_SHORT(string _Market, int _Index, int _Weekday, double _BreakOutHigh_Price, double _BreakOutLow_Price, int _BreakOutIndex)
{
   // #########################################################################################
   _1stBreakOutType       = SHORT;
   _1stEntry_Price        = 7654321;
   _1stStopLoss_Price     = 7654321;
   _1stStopLoss_R         =       0;
   _1stMaxPotential_Price = 7654321;
   _1stMaxPotential_R     =       0;
   
   _2ndBreakOutType       = NONE;
   _2ndEntry_Price        = -7654321;
   _2ndStopLoss_Price     = -7654321;
   _2ndStopLoss_R         =        0;
   _2ndMaxPotential_Price = -7654321;
   _2ndMaxPotential_R     =        0;
            
   // 1st BreakOut (POSITION_TYPE_SELL)
   _1stEntry_Price        = _BreakOutLow_Price;
   _1stStopLoss_Price     = _BreakOutHigh_Price;
   _1stStopLoss_R         = _BreakOutHigh_Price - _BreakOutLow_Price;
   
   bool _IsStoppedOut    = false;
   int  _StoppedOutIndex = 0;
   for (int i = _BreakOutIndex; i >= _Index - 48; i--)
   {    
      // Compute _MaxPotential_Price, _MaxPotenial_R
      double _Low            = iLow(   _Market, PERIOD_M15, i);
      _1stMaxPotential_Price = MathMin(_1stMaxPotential_Price, _Low);
            
      // Check if we are stopped out
      double _High = iHigh(_Market, PERIOD_M15, i);
      if (_High > _1stStopLoss_Price)
      {
         _IsStoppedOut    = true;
         _StoppedOutIndex = i;
         break;
      }
   }
       _1stMaxPotential_R  = (_1stEntry_Price - _1stMaxPotential_Price) / _1stStopLoss_R;
   if (_1stMaxPotential_R >= 1)
   {
      _1stWinner_Weekday_SHORT[0]            += 1;
      _1stWinner_Weekday_SHORT[_Weekday]     += 1;
      _1stResult_MPR_Weekday_SHORT[0]        += _1stMaxPotential_R;
      _1stResult_MPR_Weekday_SHORT[_Weekday] += _1stMaxPotential_R;
   }
   else
   {
      _1stLosser_Weekday_SHORT[0]            += 1;
      _1stLosser_Weekday_SHORT[_Weekday]     += 1;
      _1stResult_MPR_Weekday_SHORT[0]        -= 1;
      _1stResult_MPR_Weekday_SHORT[_Weekday] -= 1;
   }
   _1stMaxPotential_R_All += _1stMaxPotential_R;
   
   // #########################################################################################
   if (!_IsStoppedOut) { return; }
   
   // 2nd BreakOut (POSITION_TYPE_BUY)
   _2ndBreakOutType   = LONG;
   _2ndEntry_Price    = _BreakOutHigh_Price;
   _2ndStopLoss_Price = _BreakOutLow_Price;
   _2ndStopLoss_R     = _BreakOutHigh_Price - _BreakOutLow_Price;
   
   for (int i = _StoppedOutIndex; i > _Index - 48; i--)
   {    
      // Compute _MaxPotential_Price, _MaxPotenial_R
      double _High           = iHigh(  _Market, PERIOD_M15, i);
      _2ndMaxPotential_Price = MathMax(_2ndMaxPotential_Price, _High);
            
      // Check if we are stopped out
      double _Low = iLow(_Market, PERIOD_M15, i);
      if (_Low < _2ndStopLoss_Price) { break; }
   }
       _2ndMaxPotential_R  = (_2ndMaxPotential_Price - _2ndEntry_Price) / _2ndStopLoss_R;
   if (_2ndMaxPotential_R >= 1)
   {
      _2ndWinner_Weekday_LONG[0]            += 1;
      _2ndWinner_Weekday_LONG[_Weekday]     += 1;
      _2ndResult_MPR_Weekday_LONG[0]        += _1stMaxPotential_R;
      _2ndResult_MPR_Weekday_LONG[_Weekday] += _1stMaxPotential_R;
   }
   else
   {
      _2ndLosser_Weekday_LONG[0]            += 1;
      _2ndLosser_Weekday_LONG[_Weekday]     += 1;
      _2ndResult_MPR_Weekday_LONG[0]        -= 1;
      _2ndResult_MPR_Weekday_LONG[_Weekday] -= 1;
   }
   _2ndMaxPotential_R_All += _2ndMaxPotential_R;
}

void WriteTradeData(int _SpreadSheetHandle, string _Market, int _DayOfWeek, double _BreakOutHigh_Price, double _BreakOutLow_Price, int _BreakOutIndex)
{
   if (_1stBreakOutType == NONE) { return; }
   
   if (_2ndBreakOutType == LONG || _2ndBreakOutType == SHORT)
   {
      if (_2ndMaxPotential_Price == 7654321 || _2ndMaxPotential_Price == -7654321) { return; }
   }
   
   // Catch data
   string _M  = _Market;
   string _D  = TimeToString(iTime(_Market, PERIOD_M15, _BreakOutIndex) - PeriodSeconds(PERIOD_H1), TIME_DATE);
   string _W  = IntegerToString(   _DayOfWeek);
   string _TL = TimeToString(iTime(_Market, PERIOD_M15, _BreakOutIndex) - PeriodSeconds(PERIOD_H1), TIME_MINUTES);
   string _TS = TimeToString(iTime(_Market, PERIOD_M15, _BreakOutIndex),                            TIME_MINUTES);
         
   string _BOH = DoubleToString(_BreakOutHigh_Price, GetDigits(_Market)); StringReplace(_BOH, ".", ",");
   string _BOL = DoubleToString(_BreakOutLow_Price,  GetDigits(_Market)); StringReplace(_BOL, ".", ",");
   
   string _1stPT  = EnumToString(_1stBreakOutType);
   string _1stEP  = DoubleToString(_1stEntry_Price,                     GetDigits(_Market)); StringReplace(_1stEP,  ".", ",");
   string _1stSL  = DoubleToString(_1stStopLoss_Price,                  GetDigits(_Market)); StringReplace(_1stSL,  ".", ",");
   string _1stSLR = DoubleToString(_1stStopLoss_R * GetFactor(_Market), GetDigits(_Market)); StringReplace(_1stSLR, ".", ",");
   string _1stMP  = DoubleToString(_1stMaxPotential_Price,              GetDigits(_Market)); StringReplace(_1stMP,  ".", ",");
   string _1stMPR = DoubleToString(_1stMaxPotential_R,                  GetDigits(_Market)); StringReplace(_1stMPR, ".", ",");
   
   string _2ndPT  = EnumToString(_2ndBreakOutType);
   string _2ndEP  = "";
   string _2ndSL  = "";
   string _2ndSLR = "";
   string _2ndMP  = "";
   string _2ndMPR = "";
   if (_2ndBreakOutType != NONE)
   {
      _2ndEP  = DoubleToString(_2ndEntry_Price,                     GetDigits(_Market)); StringReplace(_2ndEP,  ".", ",");
      _2ndSL  = DoubleToString(_2ndStopLoss_Price,                  GetDigits(_Market)); StringReplace(_2ndSL,  ".", ",");
      _2ndSLR = DoubleToString(_2ndStopLoss_R * GetFactor(_Market), GetDigits(_Market)); StringReplace(_2ndSLR, ".", ",");
      _2ndMP  = DoubleToString(_2ndMaxPotential_Price,              GetDigits(_Market)); StringReplace(_2ndMP,  ".", ",");
      _2ndMPR = DoubleToString(_2ndMaxPotential_R,                  GetDigits(_Market)); StringReplace(_2ndMPR, ".", ",");
   }
   
   // Write data
   FileWrite(_SpreadSheetHandle, _M, _D, _W, _TL, _TS, _BOH, _BOL,
             _1stPT, _1stEP, _1stSL, _1stSLR, _1stMP, _1stMPR,
             _2ndPT, _2ndEP, _2ndSL, _2ndSLR, _2ndMP, _2ndMPR);
}

void WriteTradeDataSummary(int _SpreadSheetHandle, string _Market)
{
   // Write data
   if (ShowReportSingleTradeData)
   {
      string _1stMPRA = DoubleToString(_1stMaxPotential_R_All, GetDigits(_Market)); StringReplace(_1stMPRA, ".", ",");
      string _2ndMPRA = DoubleToString(_2ndMaxPotential_R_All, GetDigits(_Market)); StringReplace(_2ndMPRA, ".", ",");
      FileWrite(_SpreadSheetHandle, "", "", "", "", "", "", "",
                                    "", "", "", "", "", _1stMPRA,
                                    "", "", "", "", "", _2ndMPRA);
      FileWrite(_SpreadSheetHandle, "");
   }
   
   FileWrite(_SpreadSheetHandle, "######", "1. Summary", "(LONG)",  "######", "######", "######",
                                           "2. Summary", "(SHORT)", "######", "######", "######");
   FileWrite(_SpreadSheetHandle, "Weekday", "Winner", "Losser", "Win-Rate", "Result (T-1.0)", "Result (MPR)",
                                            "Winner", "Losser", "Win-Rate", "Result (T-1.0)", "Result (MPR)");
   for (int i = 0; i < 6; i++)
   {
      string _1stWinner_LONG  = DoubleToString(_1stWinner_Weekday_LONG[i],  0); StringReplace(_1stWinner_LONG,  ".", ",");
      string _1stLosser_LONG  = DoubleToString(_1stLosser_Weekday_LONG[i],  0); StringReplace(_1stLosser_LONG,  ".", ",");
      string _2ndWinner_SHORT = DoubleToString(_2ndWinner_Weekday_SHORT[i], 0); StringReplace(_2ndWinner_SHORT, ".", ",");
      string _2ndLosser_SHORT = DoubleToString(_2ndLosser_Weekday_SHORT[i], 0); StringReplace(_2ndLosser_SHORT, ".", ",");
      
      string _1stWinRate_LONG       = DoubleToString(_1stWinner_Weekday_LONG[i]  / (_1stWinner_Weekday_LONG[i]  + _1stLosser_Weekday_LONG[i]),  2); StringReplace(_1stWinRate_LONG, ".", ",");
      string _1stResultTarget_LONG  = DoubleToString(_1stWinner_Weekday_LONG[i]  -  _1stLosser_Weekday_LONG[i],                                 0);
      string _2ndWinRate_SHORT      = DoubleToString(_2ndWinner_Weekday_SHORT[i] / (_2ndWinner_Weekday_SHORT[i] + _2ndLosser_Weekday_SHORT[i]), 2); StringReplace(_2ndWinRate_SHORT, ".", ",");
      string _2ndResultTarget_SHORT = DoubleToString(_2ndWinner_Weekday_SHORT[i] -  _2ndLosser_Weekday_SHORT[i],                                0);
      
      string _1stMPR_LONG  = DoubleToString(_1stResult_MPR_Weekday_LONG[i],  2); StringReplace(_1stMPR_LONG,  ".", ",");
      string _2ndMPR_SHORT = DoubleToString(_2ndResult_MPR_Weekday_SHORT[i], 2); StringReplace(_2ndMPR_SHORT, ".", ",");
      
      FileWrite(_SpreadSheetHandle, GetWeekday(i), _1stWinner_LONG,  _1stLosser_LONG,  _1stWinRate_LONG,  _1stResultTarget_LONG,  _1stMPR_LONG,
                                                   _2ndWinner_SHORT, _2ndLosser_SHORT, _2ndWinRate_SHORT, _2ndResultTarget_SHORT, _2ndMPR_SHORT);
   }
     
   FileWrite(_SpreadSheetHandle, "");
   FileWrite(_SpreadSheetHandle, "######", "1. Summary", "(SHORT)", "######", "######", "######",
                                           "2. Summary", "(LONG)",  "######", "######", "######");
   FileWrite(_SpreadSheetHandle, "Weekday", "Winner", "Losser", "Win-Rate", "Result (T-1.0)", "Result (MPR)",
                                            "Winner", "Losser", "Win-Rate", "Result (T-1.0)", "Result (MPR)");
   for (int i = 0; i < 6; i++)
   {
      string _1stWinner_SHORT = DoubleToString(_1stWinner_Weekday_SHORT[i], 0); StringReplace(_1stWinner_SHORT, ".", ",");
      string _1stLosser_SHORT = DoubleToString(_1stLosser_Weekday_SHORT[i], 0); StringReplace(_1stLosser_SHORT, ".", ",");
      string _2ndWinner_LONG  = DoubleToString(_2ndWinner_Weekday_LONG[i],  0); StringReplace(_2ndWinner_LONG,  ".", ",");
      string _2ndLosser_LONG  = DoubleToString(_2ndLosser_Weekday_LONG[i],  0); StringReplace(_2ndLosser_LONG,  ".", ",");
      
      string _1stWinRate_SHORT      = DoubleToString(_1stWinner_Weekday_SHORT[i] / (_1stWinner_Weekday_SHORT[i] + _1stLosser_Weekday_SHORT[i]), 2); StringReplace(_1stWinRate_SHORT, ".", ",");
      string _1stResultTarget_SHORT = DoubleToString(_1stWinner_Weekday_SHORT[i] -  _1stLosser_Weekday_SHORT[i],                                0);
      string _2ndWinRate_LONG       = DoubleToString(_2ndWinner_Weekday_LONG[i]  / (_2ndWinner_Weekday_LONG[i]  + _2ndLosser_Weekday_LONG[i]),  2); StringReplace(_2ndWinRate_LONG, ".", ",");
      string _2ndResultTarget_LONG  = DoubleToString(_2ndWinner_Weekday_LONG[i]  -  _2ndLosser_Weekday_LONG[i],                                 0);
      
      string _1stMPR_SHORT = DoubleToString(_1stResult_MPR_Weekday_SHORT[i], 2); StringReplace(_1stMPR_SHORT, ".", ",");
      string _2ndMPR_LONG  = DoubleToString(_2ndResult_MPR_Weekday_LONG[i],  2); StringReplace(_2ndMPR_LONG,  ".", ",");
      
      FileWrite(_SpreadSheetHandle, GetWeekday(i), _1stWinner_SHORT, _1stLosser_SHORT, _1stWinRate_SHORT, _1stResultTarget_SHORT, _1stMPR_SHORT,
                                                   _2ndWinner_LONG,  _2ndLosser_LONG,  _2ndWinRate_LONG,  _2ndResultTarget_LONG,  _2ndMPR_LONG);
   }
   
}

int GetDigits(string _Market)
{
   int _D = 5; // Normal forex
   if (StringCompare(_Market,                  "XAGUSD",      true) == 0) { _D = 2; } // Silver
   if (StringCompare(_Market,                  "XAUUSD",      true) == 0) { _D = 2; } // Gold
   if (StringCompare(_Market,                  "NAT.GAS",     true) == 0) { _D = 4; } // Natural Gas
   if (StringCompare(StringSubstr(_Market, 3), "CZK",         true) == 0) { _D = 3; } // CZK
   if (StringCompare(StringSubstr(_Market, 3), "HUF",         true) == 0) { _D = 3; } // HUF
   if (StringCompare(StringSubstr(_Market, 3), "JPY",         true) == 0) { _D = 3; } // JPY
   if (StringCompare(_Market,                  ".US500Cash",  true) == 0) { _D = 2; } // S&P 500
   if (StringCompare(_Market,                  ".USTECHCash", true) == 0) { _D = 2; } // NASDAQ
   if (StringCompare(_Market,                  ".US30Cash",   true) == 0) { _D = 2; } // Dow Jones
   
   return _D;
}
   
int GetFactor(string _Market)
{
   int _Factor = 100000; // Normal forex
   if (StringCompare(_Market,                  "XAGUSD",      true) == 0) { _Factor = 100;   } // Silver
   if (StringCompare(_Market,                  "XAUUSD",      true) == 0) { _Factor = 100;   } // Gold
   if (StringCompare(_Market,                  "NAT.GAS",     true) == 0) { _Factor = 10000; } // Natural Gas
   if (StringCompare(StringSubstr(_Market, 3), "CZK",         true) == 0) { _Factor = 1000;  } // CZK
   if (StringCompare(StringSubstr(_Market, 3), "HUF",         true) == 0) { _Factor = 1000;  } // HUF
   if (StringCompare(StringSubstr(_Market, 3), "JPY",         true) == 0) { _Factor = 1000;  } // JPY
   if (StringCompare(_Market,                  ".US500Cash",  true) == 0) { _Factor = 100;   } // S&P 500
   if (StringCompare(_Market,                  ".USTECHCash", true) == 0) { _Factor = 100;   } // NASDAQ
   if (StringCompare(_Market,                  ".US30Cash",   true) == 0) { _Factor = 100;   } // Dow Jones
   
   return _Factor;
}

string GetWeekday(int _Weekday)
{
   string _D = "ALL";
   switch(_Weekday)
   {
      case 0: { _D = "ALL";       break; }
      case 1: { _D = "MONDAY";    break; }
      case 2: { _D = "TUESDAY";   break; }
      case 3: { _D = "WEDNESDAY"; break; }
      case 4: { _D = "THURSDAY";  break; }
      case 5: { _D = "FRIDAY";    break; }
   }
   
   return _D;
}

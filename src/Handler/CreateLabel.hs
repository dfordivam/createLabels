{-# LANGUAGE NoImplicitPrelude #-}
{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE TemplateHaskell #-}
{-# LANGUAGE DeriveGeneric #-}
{-# LANGUAGE GeneralizedNewtypeDeriving #-}
{-# LANGUAGE DeriveAnyClass #-}
{-# LANGUAGE MultiParamTypeClasses #-}
{-# LANGUAGE TypeFamilies #-}
module Handler.CreateLabel where

import Import
import Codec.Xlsx.Types
import Codec.Xlsx.Parser
import qualified Data.ByteString.Lazy as BSL
import qualified Data.Text as T
import qualified Data.Text.IO as T

import Control.Lens
import qualified Data.Map as Map

parseExcelAndMakeTex bs = do
  let
    fds = getDataFromExcel bs
    is = catMaybes $ Map.elems fds
  makeLatexFile is def

data ForeignData = ForeignData
  {
    numFD :: Int
    , monthFD :: Int
    , yearFD :: Int
    , qtyFD :: Int
    , nameFD :: Text
    , addr1FD :: Text
    , addr2FD :: Maybe Text
    , addr3FD :: Maybe Text
    , addr4FD :: Maybe Text
    , addr5FD :: Maybe Text
    , cityFD :: Text
    , countryFD :: Text
  }
  deriving (Show)

type FD = Map Int ForeignData

getDataFromExcel bs = Map.fromList (map (parseOneRow cm) [2..cRows])
  where
    xx = toXlsx bs
    (_,w1):_ =  _xlSheets xx
    cm = _wsCells w1
    ((cRows,_),_) = Map.findMax cm

parseOneRow cm r = (,) r $ ForeignData
  <$> (intFromCellText =<< _cellValue =<< Map.lookup (r,2) cm)
  <*> (intFromCellText =<< _cellValue =<< Map.lookup (r,3) cm)
  <*> (intFromCellText =<< _cellValue =<< Map.lookup (r,4) cm)
  <*> (intFromCellText =<< _cellValue =<< Map.lookup (r,5) cm)
  <*> (fromCellText =<< _cellValue =<< Map.lookup (r,6) cm) -- nameFD
  <*> (fromCellText =<< _cellValue =<< Map.lookup (r,7) cm)
  <*> pure (fromCellText =<< _cellValue =<< Map.lookup (r,8) cm)
  <*> pure (fromCellText =<< _cellValue =<< Map.lookup (r,9) cm)
  <*> pure (fromCellText =<< _cellValue =<< Map.lookup (r,10) cm)
  <*> pure (fromCellText =<< _cellValue =<< Map.lookup (r,11) cm)
  <*> (fromCellText =<< _cellValue =<< Map.lookup (r,12) cm) -- cityFD
  <*> (fromCellText =<< _cellValue =<< Map.lookup (r,13) cm)
  where fromCellText (CellText t) = Just t
        fromCellText _ = Nothing
        intFromCellText (CellDouble d) = Just $ floor d
        intFromCellText (CellText t) = readMay t
        intFromCellText _ = Nothing

data LatexParams = LatexParams
  {  leftPageMargin :: Int
   , rightPageMargin :: Int
   , topPageMargin :: Int
   , bottomPageMargin :: Int
   , interLabelColumn :: Int
   , interLabelRow :: Int
   , leftLabelBorder :: Int
   , rightLabelBorder :: Int
   , topLabelBorder :: Int
   , bottomLabelBorder :: Int
  }
  deriving (Show, Generic, ToJSON, FromJSON)

instance Default LatexParams where
  def = LatexParams 7 7 10 20 0 0 0 0 0 0

makeLatexFile is lp = do
  let
    header = mconcat
      [ "\\documentclass[a4paper]{article}\n"
      , "\\usepackage[newdimens]{labels}  \n"
      , "\\LabelInfotrue                  \n"
      , "\\LabelCols=2%             Number of columns of labels per page\n"
      , "\\LabelRows=8%             Number of rows of labels per page\n"
      , "\\LeftPageMargin=", (tshow $ leftPageMargin lp), "mm%      \n"
      , "\\RightPageMargin=", (tshow $ rightPageMargin lp), "mm%      \n"
      , "\\TopPageMargin=", (tshow $ topPageMargin lp), "mm%      \n"
      , "\\BottomPageMargin=", (tshow $ bottomPageMargin lp), "mm%      \n"
      , "\\InterLabelColumn=", (tshow $ interLabelColumn lp), "mm%      \n"
      , "\\InterLabelRow=", (tshow $ interLabelRow lp), "mm%      \n"
      , "\\LeftLabelBorder=", (tshow $ leftLabelBorder lp), "mm%      \n"
      , "\\RightLabelBorder=", (tshow $ rightLabelBorder lp), "mm%      \n"
      , "\\TopLabelBorder=", (tshow $ topLabelBorder lp), "mm%      \n"
      , "\\BottomLabelBorder=", (tshow $ bottomLabelBorder lp), "mm%      \n"
      , "\\begin{document}          \n"
      ]

    footer =
      "\\end{document}\n"

    item :: ForeignData -> Text
    item i = T.replace "&" "\\&" $ T.replace "#" "\\#" $ mconcat $
      [ "\\addresslabel{%\n"
      , "\\scriptsize\\sf\n"
      , "F-", (tshow $ numFD i), "/", (tshow $ monthFD i), "/", (tshow $ yearFD i), "\n"
      , "\\qquad ", (tshow $ qtyFD i), "\\\\\n"
      , "\\small\\sf\n"
      , (nameFD i), "\\\\\n"
      , (addr1FD i), "\\\\\n"]
      ++ (addr23 i) ++
      [ (tm2 (addr4FD i) (addr5FD i))
      , (cityFD i), "\\\\\n"
      , (countryFD i), "\\\\\n"
      , "}\n"
      ]
    addr23 i = case (addr4FD i) of
      Nothing -> [(tm $ addr2FD i), (tm $ addr3FD i)]
      Just _ -> [tm2 (addr2FD i) (addr3FD i)]
    tm Nothing = ""
    tm (Just t) = t <> "\\\\\n"
    tm2 Nothing _ = ""
    tm2 (Just t) Nothing = t <> "\\\\\n"
    tm2 (Just t1) (Just t2) = t1 <>", " <> t2 <>"\\\\\n"

    content = mconcat $ header : ((map item is) ++ [footer])
  T.writeFile "out.tex" content

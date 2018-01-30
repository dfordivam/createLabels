{-# LANGUAGE NoImplicitPrelude #-}
{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE TemplateHaskell #-}
{-# LANGUAGE MultiParamTypeClasses #-}
{-# LANGUAGE TypeFamilies #-}
module Handler.Home where

import Import
import Yesod.Form.Bootstrap3 (BootstrapFormLayout (..), renderBootstrap3)
import Text.Julius (RawJS (..))
import Data.Conduit.List as Conduit
import Handler.CreateLabel
import qualified Data.ByteString as BS
import Yesod.Core.Types

-- Define our data that will be used for creating the form.
data FileForm = FileForm
    { fileInfo :: FileInfo
    , latexParams :: LatexParams
    }

-- This is a handler function for the GET request method on the HomeR
-- resource pattern. All of your resource patterns are defined in
-- config/routes
--
-- The majority of the code you will write in Yesod lives in these handler
-- functions. You can spread them across multiple files if you are so
-- inclined, or create a single monolithic file.
getHomeR :: Handler Html
getHomeR = do
    (formWidget, formEnctype) <- generateFormPost (sampleForm def)
    let submission = Nothing :: Maybe FileForm
        handlerName = "getHomeR" :: Text
    defaultLayout $ do
        let (commentFormId, commentTextareaId, commentListId) = commentIds
        aDomId <- newIdent
        setTitle "Welcome To Yesod!"
        $(widgetFile "homepage")

postHomeR :: Handler Html
postHomeR = do
    ((result, formWidget), formEnctype) <- runFormPost (sampleForm def)
    let handlerName = "postHomeR" :: Text
        submission = case result of
            FormSuccess res -> Just res
            _ -> Nothing

    case submission of
      (Just (FileForm fi lp)) -> do
        bs <- liftIO $ do
          (fileMove fi) "input.xlsx"
          BS.readFile "input.xlsx"
        res <- liftIO $ parseExcelAndMakePdf bs lp
        case res of
          (Left badRows) -> do
            setMessage (toHtml $ "Cannot Parse rows: " <> tshow badRows)
            redirect HomeR
          (Right _) -> do
          sendFile "application/pdf" "out.pdf"
      Nothing -> do
        setMessage ("Error uploading File")
        redirect HomeR

sampleForm :: LatexParams -> Form FileForm
sampleForm lp = renderBootstrap3 BootstrapBasicForm $ FileForm
    <$> fileAFormReq "Choose a file"
    <*> (LatexParams
         <$> areq intField ( "Left Page Margin") (Just $ leftPageMargin lp)
         <*> areq intField ( "Right Page Margin") (Just $ rightPageMargin lp)
         <*> areq intField ( "Top Page Margin") (Just $ topPageMargin lp)
         <*> areq intField ( "Bottom Page Margin") (Just $ bottomPageMargin lp)
         <*> areq intField ( "Inter Colomn Margin") (Just $ interLabelColumn lp)
         <*> areq intField ( "Inter Row Margin") (Just $ interLabelRow lp)
         <*> areq intField ( "Left Label Margin") (Just $ leftLabelBorder lp)
         <*> areq intField ( "Right Label Margin") (Just $ rightLabelBorder lp)
         <*> areq intField ( "Top Label Margin") (Just $ topLabelBorder lp)
         <*> areq intField ( "Bottom Label Margin") (Just $ bottomLabelBorder lp))

commentIds :: (Text, Text, Text)
commentIds = ("js-commentForm", "js-createCommentTextarea", "js-commentList")

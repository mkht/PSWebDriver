using System;
using System.Drawing;
using System.Collections.Generic;
using AnimatedGif;

// //////////////////////////////////////////
// Summary: PSWebDriverからAnimatedGifライブラリを使うためのラッパークラス
// Author : mkht
// Date   : 2018/01/01
// License: Copyright (c) 2018 mkht
//          Released under the MIT license
//          http://opensource.org/licenses/mit-license.php
// //////////////////////////////////////////

/// <summary>
/// PSWebDriverからAnimatedGifライブラリを使うためのラッパークラス
/// </summary>
/// <seealso> https://github.com/mrousavy/AnimatedGif </seealso>
public class AnimatedGifWrapper
{
    private List<Image> ImageList;  //Image buffer
    private ImageConverter Converter;

    // Constructor
    public AnimatedGifWrapper() {
        this.ImageList = new List<Image>();
        this.Converter = new ImageConverter();
    }

    /// <summary> Gifアニメーションに画像フレームを追加
    /// <param name="image">画像フレーム</param>
    public void AddFrame(Image image) {
        this.ImageList.Add(image);
    }

    /// <summary> Gifアニメーションに画像フレームを追加(Byte[])
    /// <param name="image">画像フレーム(Byte[])</param>
    public void AddFrame(Byte[] byteArray) {
        try{
            this.ImageList.Add((Image)this.Converter.ConvertFrom(byteArray));
        }
        catch {
        }
    }

    /// <summary> Gifアニメーションをファイルに保存
    /// <param name="savePath">保存先ファイル名</param>
    /// <param name="delay">フレームディレイ(ms)</param>
    public void Save(string savePath, int delay){
        using (AnimatedGifCreator gifCreator = AnimatedGif.AnimatedGif.Create(savePath, delay, 1)) { //NoRepeat
            foreach (Image img in this.ImageList) {
                using (img) {
                    gifCreator.AddFrame(img, GifQuality.Default);
                }
            }
        }
        //Clear Image buffer
        this.ImageList.Clear();
        this.ImageList.TrimExcess();
    }

    // Dispose
    public void Dispose(){
        this.Converter = null;
        this.ImageList.Clear();
        this.ImageList.TrimExcess();
    }
}
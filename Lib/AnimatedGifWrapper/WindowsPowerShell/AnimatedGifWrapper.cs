using System;
using System.IO;
using System.IO.Compression;
using System.Drawing;
using System.Drawing.Imaging;
using System.Collections.Generic;
using AnimatedGif;

// //////////////////////////////////////////
// Summary: PSWebDriverからAnimatedGifライブラリを使うためのラッパークラス
// Author : mkht
// Date   : 2020/06/07
// License: Copyright (c) 2020 mkht
//          Released under the MIT license
//          http://opensource.org/licenses/mit-license.php
// //////////////////////////////////////////

/// <summary>
/// PSWebDriverからAnimatedGifライブラリを使うためのラッパークラス
/// </summary>
/// <seealso> https://github.com/mrousavy/AnimatedGif </seealso>
public class AnimatedGifWrapper
{
    private List<MemoryStream> ImageList;  //Image buffer (Compressed)
    private ImageConverter Converter;

    // Constructor
    public AnimatedGifWrapper() {
        this.ImageList = new List<MemoryStream>();
        this.Converter = new ImageConverter();
    }

    /// <summary> Gifアニメーションに画像フレームを追加
    /// <param name="image">画像フレーム</param>
    public void AddFrame(Image image) {
        this.ImageList.Add(Compress(image));    //Bitmapのまま格納するとメモリを食い潰すので圧縮保持する
    }

    /// <summary> Gifアニメーションに画像フレームを追加(byte[])
    /// <param name="image">画像フレーム(byte[])</param>
    public void AddFrame(byte[] byteArray) {
        try{
            this.ImageList.Add(Compress((Image)this.Converter.ConvertFrom(byteArray)));
        }
        catch {
        }
    }

    /// <summary> Gifアニメーションをファイルに保存
    /// <param name="savePath">保存先ファイル名</param>
    /// <param name="delay">フレームディレイ(ms)</param>
    public void Save(string savePath, int delay){
        using (AnimatedGifCreator gifCreator = AnimatedGif.AnimatedGif.Create(savePath, delay, 1)) { //NoRepeat
            foreach (MemoryStream frame in this.ImageList) {
                using (Image img = (Image)Decompress(frame)){
                    gifCreator.AddFrame(img, delay, GifQuality.Default);
                    frame.Dispose();
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

    // compress image
    private static MemoryStream Compress(Image image)
    {
        MemoryStream output = new MemoryStream();
        image.Save(output, ImageFormat.Gif);
        return output;
    }

    // decompress image
    private static Image Decompress(MemoryStream stream)
    {
        stream.Seek(0, SeekOrigin.Begin);
        return new Bitmap(stream);
    }
}
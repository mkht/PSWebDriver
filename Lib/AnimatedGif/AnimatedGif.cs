using System;
using System.IO;
using System.Windows.Media;
using System.Windows.Media.Imaging;

public class AnimatedGif
{
    private GifBitmapEncoder encoder;

    //Constructor
    public AnimatedGif() {
        this.encoder = new GifBitmapEncoder();
    }

    public void AddFrame(BitmapFrame bmpFrame) {
        this.encoder.Frames.Add(bmpFrame);
    }

    public void AddFrame(Byte[] byteArray) {
        try{
            BitmapSource bmpSource =  (BitmapSource)new ImageSourceConverter().ConvertFrom(byteArray);
            this.encoder.Frames.Add((BitmapFrame)BitmapFrame.Create(bmpSource));
        }
        catch {
        }
    }

    public void Save(string savePath){
        using (FileStream outputFileStream = new FileStream(savePath, FileMode.Create, FileAccess.Write, FileShare.None)){
            encoder.Save(outputFileStream);
        }
    }

    public void Dispose(){
        this.encoder = null;
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
    }
}
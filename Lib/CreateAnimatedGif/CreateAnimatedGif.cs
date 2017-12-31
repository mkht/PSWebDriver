using System.IO;
using System.Windows.Media.Imaging;

public class CreateAnimatedGif
{
    private GifBitmapEncoder encoder;

    //Constructor
    public CreateAnimatedGif() {
        this.encoder = new GifBitmapEncoder();
    }

    public void AddFrame(BitmapFrame bmpFrame) {
        this.encoder.Frames.Add(bmpFrame);
    }

    public void Save(string savePath){
        using (FileStream outputFileStream = new FileStream(savePath, FileMode.Create, FileAccess.Write, FileShare.None)){
            encoder.Save(outputFileStream);
        }
    }
}
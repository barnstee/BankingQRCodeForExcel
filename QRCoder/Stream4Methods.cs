// original code from https://github.com/codebude/QRCoder
// by Raffael Herrmann and was first released in 10/2013. It's licensed under the MIT license.

namespace QRCoder.Framework4._0Methods
{
    class Stream4Methods
    {
        public static void CopyTo(System.IO.Stream input, System.IO.Stream output)
        {
            byte[] buffer = new byte[16 * 1024];
            int bytesRead;
            while ((bytesRead = input.Read(buffer, 0, buffer.Length)) > 0)
            {
                output.Write(buffer, 0, bytesRead);
            }
        }
    }
}

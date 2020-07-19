// original code from https://github.com/codebude/QRCoder
// by Raffael Herrmann and was first released in 10/2013. It's licensed under the MIT license.

using System;

namespace QRCoder.Exceptions
{
    public class DataTooLongException : Exception
    {
        public DataTooLongException(string eccLevel, string encodingMode, int maxSizeByte) : base(
            $"The given payload exceeds the maximum size of the QR code standard. The maximum size allowed for the choosen paramters (ECC level={eccLevel}, EncodingMode={encodingMode}) is {maxSizeByte} byte."
        ){}
    }
}

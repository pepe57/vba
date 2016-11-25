using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VbaReader.Compression
{
    public class CompressedChunkHeader
    {
        public readonly ushort CompressedChunkSize;
        public readonly Byte CompressedChunkFlag;
        protected readonly Byte CompressedChunkSignature;

        protected readonly ushort Header; // 2-byte / unsigned 16 bit
  

        public CompressedChunkHeader(XlBinaryReader Data)
        {
            // Algorithm as per page 60
            this.Header = Data.ReadUInt16();
  
            Console.WriteLine("CompressionChunkHeader Data Bytes: {0}  (uint16: {1})", this.Header.ToBitString(), this.Header);

            this.CompressedChunkFlag = ExtractCompressionChunkFlag();
            this.CompressedChunkSize = ExtractCompressionChunkSize();
            this.CompressedChunkSignature = ExtractCompressionChunkSignature();

            Validate();

        }

        protected void Validate()
        {
            ValidateSignature();
            ValidateFlag();
            ValidateSize();
        }

        /// <summary>
        /// Validates this.CompressedChunkFlag
        /// </summary>
        protected void ValidateFlag()
        {
            if(this.CompressedChunkFlag != 0x00 && this.CompressedChunkFlag != 0x01)
            {
                throw new FormatException(String.Format("Expected CompressedChunkFlag to be either 0x00 or 0x01, but was {0:X}", this.CompressedChunkFlag));
            }
        }

        /// <summary>
        /// Validates this.CompressedChunkSignature
        /// </summary>
        protected void ValidateSignature()
        {
            if (this.CompressedChunkSignature != (Byte)0x03)
            {
                throw new FormatException(String.Format("Signature byte expected 0x03, but was 0x{0:X} (binary: {1}). (Header: 0b{2})", this.CompressedChunkSignature, this.CompressedChunkSignature.ToBitString(), Header.ToBitString()));
            }
        }

        /// <summary>
        /// Validates this.CompressedChunkSize
        /// </summary>
        protected void ValidateSize()
        {
            if (this.CompressedChunkFlag == 0x00 && this.CompressedChunkSize != 4095)
            {
                throw new FormatException(String.Format("CompressionChunkFlag = {0:X}, expected CompressedChunkSize 4095, but was {1}", this.CompressedChunkFlag, this.CompressedChunkSize));
            }

            if(this.CompressedChunkSize < 0 || (this.CompressedChunkFlag == 0x01 && this.CompressedChunkSize > 4095))
            {
                throw new FormatException(String.Format("Expected CompressedChunkSize to be between 0 and 4095, but was {0}", this.CompressedChunkSize));
            }
        }

        protected ushort ExtractCompressionChunkSize()
        {
            if (this.CompressedChunkFlag == (Byte)0x00)
                return 4095;

            // Extract CompressionChunkSize
            // page 66
            Console.WriteLine("Temp value: 0x{0:X}", this.Header);
            var temp = (ushort)(this.Header & (ushort)0x0FFF);
            temp = (ushort)(temp + (ushort)3);
            ushort size = temp;

            if (size < 3)
                throw new FormatException("Size was < 3");

            if (size > 4098)
                throw new FormatException("Size was > 4098");

            return size;
        }

        protected Byte ExtractCompressionChunkFlag()
        {
            // Extract CompressionChunkFlag
            // page 67
            Console.WriteLine("Extracting CompressionChunkFlag from header {0}", this.Header);
            var temp = this.Header & 0x8000;
            temp = temp >> 15;  // right shift 15 bits
            Byte result = (Byte)temp;
            Console.WriteLine("Extracted flag: {0} (cast to Byte: {1})", temp, result);

            return result;
        }

        protected Byte ExtractCompressionChunkSignature()
        {
            ushort signature = (ushort)(Header >> 12);
            signature = (ushort)(signature & 0x07);

            var result = (byte)signature;

            return (Byte)signature;
        }



    }
}

using System;
using System.Collections.Generic;
using System.Text;

// MS-OVBA decompression + VBA "dir" stream parsing.
// Algorithm ported from pyOpenVBA/vba.py (MIT).

class ByteBuf
{
    byte[] buf;
    int len;

    public ByteBuf(int capacity)
    {
        if (capacity < 16) capacity = 16;
        buf = new byte[capacity];
        len = 0;
    }

    public int Length { get { return len; } }

    public byte this[int index] { get { return buf[index]; } }

    void Ensure(int extra)
    {
        int need = len + extra;
        if (need <= buf.Length) return;
        int cap = buf.Length;
        while (cap < need) cap *= 2;
        byte[] nb = new byte[cap];
        Buffer.BlockCopy(buf, 0, nb, 0, len);
        buf = nb;
    }

    public void Add(byte b)
    {
        Ensure(1);
        buf[len++] = b;
    }

    public void AddRange(byte[] src, int off, int count)
    {
        Ensure(count);
        Buffer.BlockCopy(src, off, buf, len, count);
        len += count;
    }

    // Copy from already-produced output. BlockCopy is safe here only when
    // the ranges do not overlap (offset >= length); overlapping runs must be
    // replicated byte by byte.
    public void AddSelfRange(int src, int count)
    {
        Ensure(count);
        Buffer.BlockCopy(buf, src, buf, len, count);
        len += count;
    }

    public void AddSelf(int src, int count)
    {
        Ensure(count);
        for (int i = 0; i < count; i++) buf[len++] = buf[src + i];
    }

    public byte[] ToArray()
    {
        byte[] r = new byte[len];
        Buffer.BlockCopy(buf, 0, r, 0, len);
        return r;
    }
}

class VbaModule
{
    public string Name = "";
    public string NameUnicode = "";
    public string StreamName = "";
    public string StreamNameUnicode = "";
    public int TextOffset = 0;
    public int Kind = 0;          // 0 = standard, 1 = class/other
    public bool IsReadOnly = false;
    public bool IsPrivate = false;
    public string Code = "";
}

class VbaDirInfo
{
    public int CodePage = 1252;
    public List<VbaModule> Modules = new List<VbaModule>();
}

static class Ovba
{
    // ------------------------------------------------------------ decompress

    public static byte[] Decompress(byte[] data)
    {
        if (data == null || data.Length == 0 || data[0] != 0x01)
            throw new Exception("MS-OVBA: missing 0x01 signature");

        int pos = 1;
        int n = data.Length;
        ByteBuf outp = new ByteBuf(n * 3);

        while (pos < n)
        {
            if (pos + 2 > n) throw new Exception("MS-OVBA: truncated chunk header");

            int header = data[pos] | (data[pos + 1] << 8);
            int chunkDataSize = (header & 0x0FFF) + 1;
            int signature = (header >> 12) & 0x7;
            int compressed = (header >> 15) & 0x1;
            int headerOffset = pos;
            pos += 2;

            if (signature != 3)
                throw new Exception("MS-OVBA: bad chunk signature at " + headerOffset);

            int chunkEnd = pos + chunkDataSize;
            if (chunkEnd > n)
                throw new Exception("MS-OVBA: truncated chunk at " + headerOffset);

            int chunkStart = outp.Length;

            if (compressed == 0)
            {
                if (chunkDataSize != 4096)
                    throw new Exception("MS-OVBA: raw chunk is not 4096 bytes");
                outp.AddRange(data, pos, 4096);
                pos += 4096;
            }
            else
            {
                while (pos < chunkEnd)
                {
                    int flagByte = data[pos++];
                    int bit = 0;

                    while (bit < 8)
                    {
                        if (pos >= chunkEnd || pos >= n) break;

                        if (((flagByte >> bit) & 1) == 1)
                        {
                            if (pos + 2 > n) throw new Exception("MS-OVBA: truncated copy token");
                            int token = data[pos] | (data[pos + 1] << 8);
                            pos += 2;

                            int difference = outp.Length - chunkStart;
                            if (difference < 1) difference = 1;
                            int bitCount = CeilLog2(difference);
                            if (bitCount < 4) bitCount = 4;

                            int lengthMask = 0xFFFF >> bitCount;
                            int offsetMask = (~lengthMask) & 0xFFFF;

                            int length = (token & lengthMask) + 3;
                            int offset = ((token & offsetMask) >> (16 - bitCount)) + 1;

                            int copySrc = outp.Length - offset;
                            if (copySrc < 0 || copySrc < chunkStart)
                                throw new Exception("MS-OVBA: copy token out of range");

                            if (offset >= length)
                                outp.AddSelfRange(copySrc, length);
                            else
                                outp.AddSelf(copySrc, length);

                            bit++;
                        }
                        else
                        {
                            int runStart = pos;
                            while (bit < 8 && ((flagByte >> bit) & 1) == 0 && pos < chunkEnd && pos < n)
                            {
                                pos++;
                                bit++;
                            }
                            outp.AddRange(data, runStart, pos - runStart);
                        }
                    }
                }
            }
        }

        return outp.ToArray();
    }

    static int CeilLog2(int n)
    {
        int r = 0;
        n -= 1;
        while (n > 0) { n >>= 1; r++; }
        return r;
    }

    // ------------------------------------------------------------ dir stream

    static int ReadU16(byte[] b, ref int pos)
    {
        if (pos + 2 > b.Length) throw new Exception("dir: truncated u16");
        int v = b[pos] | (b[pos + 1] << 8);
        pos += 2;
        return v;
    }

    static uint ReadU32(byte[] b, ref int pos)
    {
        if (pos + 4 > b.Length) throw new Exception("dir: truncated u32");
        uint v = (uint)(b[pos] | (b[pos + 1] << 8) | (b[pos + 2] << 16) | (b[pos + 3] << 24));
        pos += 4;
        return v;
    }

    static string Decode(byte[] b, int off, int len, int codePage)
    {
        if (len <= 0) return "";
        try { return Encoding.GetEncoding(codePage).GetString(b, off, len); }
        catch { return Encoding.GetEncoding(28591).GetString(b, off, len); }
    }

    public static VbaDirInfo ParseDir(byte[] raw)
    {
        VbaDirInfo info = new VbaDirInfo();
        VbaModule cur = null;
        int pos = 0;

        while (pos + 2 <= raw.Length)
        {
            int id = ReadU16(raw, ref pos);

            if (id == 0x0010) // dir terminator
            {
                if (pos + 4 <= raw.Length) pos += 4;
                break;
            }

            if (id == 0x0009) // PROJECTVERSION: no size field
            {
                if (pos + 10 > raw.Length) break;
                pos += 4; // reserved
                pos += 4; // major
                pos += 2; // minor
                continue;
            }

            if (id == 0x0019) // MODULENAME starts a new module
            {
                if (cur != null) info.Modules.Add(cur);
                cur = new VbaModule();
                uint size = ReadU32(raw, ref pos);
                if (pos + (int)size > raw.Length) break;
                cur.Name = Decode(raw, pos, (int)size, info.CodePage);
                pos += (int)size;
                continue;
            }

            if (pos + 4 > raw.Length) break;
            uint sz = ReadU32(raw, ref pos);
            if (pos + (int)sz > raw.Length) break;
            int dataStart = pos;
            pos += (int)sz;

            switch (id)
            {
                case 0x0003: // PROJECTCODEPAGE
                    if (sz >= 2) info.CodePage = raw[dataStart] | (raw[dataStart + 1] << 8);
                    break;

                case 0x0047: // MODULENAMEUNICODE
                    if (cur != null) cur.NameUnicode = Encoding.Unicode.GetString(raw, dataStart, (int)sz);
                    break;

                case 0x001A: // MODULESTREAMNAME
                    if (cur != null) cur.StreamName = Decode(raw, dataStart, (int)sz, info.CodePage);
                    break;

                case 0x0032: // MODULESTREAMNAME unicode
                    if (cur != null) cur.StreamNameUnicode = Encoding.Unicode.GetString(raw, dataStart, (int)sz);
                    break;

                case 0x0031: // MODULEOFFSET
                    if (cur != null && sz >= 4) cur.TextOffset = (int)BitConverter.ToUInt32(raw, dataStart);
                    break;

                case 0x0021: if (cur != null) cur.Kind = 0; break; // MODULETYPE standard
                case 0x0022: if (cur != null) cur.Kind = 1; break; // MODULETYPE other
                case 0x0025: if (cur != null) cur.IsReadOnly = true; break;
                case 0x0028: if (cur != null) cur.IsPrivate = true; break;

                case 0x002B: // module terminator
                    if (cur != null) { info.Modules.Add(cur); cur = null; }
                    break;
            }
        }

        if (cur != null && cur.Name.Length > 0) info.Modules.Add(cur);

        // Prefer the Unicode records when present (they are authoritative).
        for (int i = 0; i < info.Modules.Count; i++)
        {
            VbaModule m = info.Modules[i];
            if (m.NameUnicode.Length > 0) m.Name = m.NameUnicode;
            if (m.StreamNameUnicode.Length > 0) m.StreamName = m.StreamNameUnicode;
            if (m.StreamName.Length == 0) m.StreamName = m.Name;
        }

        return info;
    }

    // ------------------------------------------------------------ project

    public static List<VbaModule> ReadProject(byte[] vbaProjectBin, out int codePage)
    {
        Cfb cfb = new Cfb(vbaProjectBin);

        byte[] dirCompressed;
        try { dirCompressed = cfb.GetStreamInStorage("VBA", "dir"); }
        catch { dirCompressed = cfb.GetStream("dir"); }

        byte[] dirRaw = Decompress(dirCompressed);
        VbaDirInfo info = ParseDir(dirRaw);
        codePage = info.CodePage;

        for (int i = 0; i < info.Modules.Count; i++)
        {
            VbaModule m = info.Modules[i];
            byte[] streamCompressed;
            try { streamCompressed = cfb.GetStreamInStorage("VBA", m.StreamName); }
            catch { streamCompressed = cfb.GetStream(m.StreamName); }

            if (m.TextOffset < 0 || m.TextOffset > streamCompressed.Length)
                throw new Exception("MODULEOFFSET out of range for " + m.Name);

            byte[] src = new byte[streamCompressed.Length - m.TextOffset];
            Buffer.BlockCopy(streamCompressed, m.TextOffset, src, 0, src.Length);

            byte[] raw = Decompress(src);
            m.Code = Decode(raw, 0, raw.Length, info.CodePage);
        }

        return info.Modules;
    }
}

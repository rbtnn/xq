using System;
using System.Collections.Generic;
using System.Text;

// Minimal MS-CFB (Compound File Binary) reader, enough to pull VBA streams
// out of xl/vbaProject.bin. Ported from pyOpenVBA/cfb.py (MIT).
class CfbEntry
{
    public string Name;
    public int Type;      // 1 = storage, 2 = stream, 5 = root
    public int Left;
    public int Right;
    public int Child;
    public long Start;
    public long Size;
}

class Cfb
{
    const uint FREESECT = 0xFFFFFFFF;
    const uint ENDOFCHAIN = 0xFFFFFFFE;
    const uint FATSECT = 0xFFFFFFFD;
    const uint DIFSECT = 0xFFFFFFFC;

    byte[] data;
    int majorVer;
    int sectorSize;
    int miniSectorSize;
    uint miniCutoff;
    uint[] fat;
    uint[] minifat;
    List<CfbEntry> dir;
    byte[] miniStream;

    public Cfb(byte[] bytes)
    {
        data = bytes;
        Parse();
    }

    // ---------------------------------------------------------------- parse

    void Parse()
    {
        if (data.Length < 512) throw new Exception("CFB: file too small");
        if (data[0] != 0xD0 || data[1] != 0xCF || data[2] != 0x11 || data[3] != 0xE0 ||
            data[4] != 0xA1 || data[5] != 0xB1 || data[6] != 0x1A || data[7] != 0xE1)
            throw new Exception("CFB: bad magic");

        majorVer = BitConverter.ToUInt16(data, 26);
        int sectorShift = BitConverter.ToUInt16(data, 30);
        int miniShift = BitConverter.ToUInt16(data, 32);

        if (majorVer != 3 && majorVer != 4) throw new Exception("CFB: unsupported major version " + majorVer);
        if (sectorShift != (majorVer == 3 ? 9 : 12) || miniShift != 6)
            throw new Exception("CFB: unsupported sector sizes");

        sectorSize = 1 << sectorShift;
        miniSectorSize = 1 << miniShift;

        uint firstDir = BitConverter.ToUInt32(data, 48);
        miniCutoff = BitConverter.ToUInt32(data, 56);
        uint firstMiniFat = BitConverter.ToUInt32(data, 60);
        uint firstDifat = BitConverter.ToUInt32(data, 68);
        uint numDifat = BitConverter.ToUInt32(data, 72);

        List<uint> difat = new List<uint>();
        for (int i = 0; i < 109; i++) difat.Add(BitConverter.ToUInt32(data, 76 + i * 4));

        if (firstDifat != ENDOFCHAIN && firstDifat != FREESECT && numDifat > 0)
        {
            uint sect = firstDifat;
            int per = sectorSize / 4 - 1;
            for (uint k = 0; k < numDifat && sect != ENDOFCHAIN && sect != FREESECT; k++)
            {
                byte[] sd = Sector(sect);
                for (int i = 0; i < per; i++) difat.Add(BitConverter.ToUInt32(sd, i * 4));
                sect = BitConverter.ToUInt32(sd, sectorSize - 4);
            }
        }

        // FAT
        byte[] fatBytes = ConcatSectors(difat, FATSECT, DIFSECT);
        fat = new uint[fatBytes.Length / 4];
        for (int i = 0; i < fat.Length; i++) fat[i] = BitConverter.ToUInt32(fatBytes, i * 4);

        // mini FAT
        if (firstMiniFat != ENDOFCHAIN && firstMiniFat != FREESECT)
        {
            byte[] mfBytes = Concat(Chain(firstMiniFat));
            minifat = new uint[mfBytes.Length / 4];
            for (int i = 0; i < minifat.Length; i++) minifat[i] = BitConverter.ToUInt32(mfBytes, i * 4);
        }
        else
        {
            minifat = new uint[0];
        }

        // directory
        byte[] dirRaw = Concat(Chain(firstDir));
        dir = new List<CfbEntry>();
        for (int off = 0; off + 128 <= dirRaw.Length; off += 128)
            dir.Add(ParseEntry(dirRaw, off));

        // mini stream (root entry)
        if (dir.Count > 0 && dir[0].Start != ENDOFCHAIN && dir[0].Start != FREESECT)
            miniStream = Concat(Chain((uint)dir[0].Start));
        else
            miniStream = new byte[0];
    }

    byte[] ConcatSectors(List<uint> sectors, uint stop1, uint stop2)
    {
        int count = 0;
        for (int i = 0; i < sectors.Count; i++)
        {
            uint s = sectors[i];
            if (s == FREESECT || s == ENDOFCHAIN || s == stop1 || s == stop2) break;
            count++;
        }
        byte[] outp = new byte[count * sectorSize];
        for (int i = 0; i < count; i++)
        {
            byte[] sd = Sector(sectors[i]);
            Buffer.BlockCopy(sd, 0, outp, i * sectorSize, sectorSize);
        }
        return outp;
    }

    byte[] Concat(List<uint> sectors)
    {
        byte[] outp = new byte[sectors.Count * sectorSize];
        for (int i = 0; i < sectors.Count; i++)
        {
            byte[] sd = Sector(sectors[i]);
            Buffer.BlockCopy(sd, 0, outp, i * sectorSize, sectorSize);
        }
        return outp;
    }

    byte[] Sector(uint index)
    {
        long off = ((long)index + 1) * sectorSize;
        if (off < 0 || off + sectorSize > data.Length)
            throw new Exception("CFB: sector out of range: " + index);
        byte[] s = new byte[sectorSize];
        Buffer.BlockCopy(data, (int)off, s, 0, sectorSize);
        return s;
    }

    List<uint> Chain(uint start)
    {
        List<uint> list = new List<uint>();
        HashSet<uint> seen = new HashSet<uint>();
        uint cur = start;
        while (cur != ENDOFCHAIN && cur != FREESECT)
        {
            if (cur >= fat.Length || !seen.Add(cur))
                throw new Exception("CFB: bad FAT chain at " + cur);
            list.Add(cur);
            cur = fat[cur];
        }
        return list;
    }

    List<uint> MiniChain(uint start)
    {
        List<uint> list = new List<uint>();
        HashSet<uint> seen = new HashSet<uint>();
        uint cur = start;
        while (cur != ENDOFCHAIN && cur != FREESECT)
        {
            if (cur >= minifat.Length || !seen.Add(cur))
                throw new Exception("CFB: bad mini-FAT chain at " + cur);
            list.Add(cur);
            cur = minifat[cur];
        }
        return list;
    }

    CfbEntry ParseEntry(byte[] raw, int off)
    {
        CfbEntry e = new CfbEntry();
        int nameLen = BitConverter.ToUInt16(raw, off + 64);
        int chars = nameLen / 2 - 1;
        if (chars < 0) chars = 0;
        e.Name = Encoding.Unicode.GetString(raw, off, chars * 2);
        e.Type = raw[off + 66];
        e.Left = (int)BitConverter.ToUInt32(raw, off + 68);
        e.Right = (int)BitConverter.ToUInt32(raw, off + 72);
        e.Child = (int)BitConverter.ToUInt32(raw, off + 76);

        uint startLow = BitConverter.ToUInt32(raw, off + 116);
        uint sizeLow = BitConverter.ToUInt32(raw, off + 120);
        uint sizeHigh = BitConverter.ToUInt32(raw, off + 124);

        e.Start = startLow;
        e.Size = (majorVer == 4) ? (((long)sizeHigh << 32) | sizeLow) : sizeLow;
        return e;
    }

    // ---------------------------------------------------------------- read

    public byte[] ReadStream(CfbEntry e)
    {
        if (e.Size == 0) return new byte[0];

        if (e.Size < miniCutoff && e.Type != 5)
        {
            List<uint> chain = MiniChain((uint)e.Start);
            byte[] raw = new byte[chain.Count * miniSectorSize];
            for (int i = 0; i < chain.Count; i++)
            {
                int off = (int)chain[i] * miniSectorSize;
                if (off + miniSectorSize > miniStream.Length)
                    throw new Exception("CFB: mini sector out of range");
                Buffer.BlockCopy(miniStream, off, raw, i * miniSectorSize, miniSectorSize);
            }
            byte[] res = new byte[e.Size];
            Buffer.BlockCopy(raw, 0, res, 0, (int)e.Size);
            return res;
        }

        byte[] big = Concat(Chain((uint)e.Start));
        byte[] result = new byte[e.Size];
        Buffer.BlockCopy(big, 0, result, 0, (int)e.Size);
        return result;
    }

    int FindStorage(string name)
    {
        for (int i = 0; i < dir.Count; i++)
        {
            if (dir[i].Type == 1 && string.Equals(dir[i].Name, name, StringComparison.OrdinalIgnoreCase))
                return i;
        }
        return -1;
    }

    List<int> Subtree(int rootId)
    {
        List<int> outp = new List<int>();
        HashSet<int> seen = new HashSet<int>();
        Stack<int> stack = new Stack<int>();
        if (rootId >= 0 && rootId < dir.Count) stack.Push(rootId);

        while (stack.Count > 0)
        {
            int node = stack.Pop();
            if (node < 0 || node >= dir.Count || !seen.Add(node)) continue;
            outp.Add(node);
            CfbEntry e = dir[node];
            if (e.Left >= 0) stack.Push(e.Left);
            if (e.Right >= 0) stack.Push(e.Right);
        }
        return outp;
    }

    public byte[] GetStream(string name)
    {
        for (int i = 0; i < dir.Count; i++)
        {
            if (dir[i].Type == 2 && string.Equals(dir[i].Name, name, StringComparison.OrdinalIgnoreCase))
                return ReadStream(dir[i]);
        }
        throw new Exception("CFB: stream not found: " + name);
    }

    public byte[] GetStreamInStorage(string storage, string name)
    {
        int parent = FindStorage(storage);
        if (parent < 0) throw new Exception("CFB: storage not found: " + storage);

        List<int> children = Subtree(dir[parent].Child);
        for (int i = 0; i < children.Count; i++)
        {
            CfbEntry e = dir[children[i]];
            if (e.Type == 2 && string.Equals(e.Name, name, StringComparison.OrdinalIgnoreCase))
                return ReadStream(e);
        }
        throw new Exception("CFB: stream not found: " + storage + "/" + name);
    }
}

package ole2

import (
	"encoding/binary"
	"fmt"
	"io"
)

var ENDOFCHAIN = uint32(0xFFFFFFFE) //-2
var FREESECT = uint32(0xFFFFFFFF)   // -1

type Ole struct {
	header   *Header
	Lsector  uint32
	Lssector uint32
	SecID    []uint32
	SSecID   []uint32
	Files    []File
	reader   io.ReadSeeker
}

func Open(reader io.ReadSeeker, charset string) (ole *Ole, err error) {
	var header *Header
	var hbts = make([]byte, 512)
	reader.Read(hbts)
	if header, err = parseHeader(hbts); err == nil {
		ole = new(Ole)
		ole.reader = reader
		ole.header = header
		ole.Lsector = 512 //TODO
		ole.Lssector = 64 //TODO
		err = ole.readMSAT()
		return ole, err
	}

	return nil, err
}

func (o *Ole) ListDir() (dir []*File, err error) {
	sector := o.stream_read(o.header.Dirstart, 0)
	dir = make([]*File, 0)
	for {
		d := new(File)
		err = binary.Read(sector, binary.LittleEndian, d)
		if err == nil && d.Type != EMPTY {
			dir = append(dir, d)
		} else {
			break
		}
	}
	if err == io.EOF && dir != nil {
		return dir, nil
	}

	return
}

func (o *Ole) OpenFile(file *File, root *File) io.ReadSeeker {
	if file.Size < o.header.Sectorcutoff {
		return o.short_stream_read(file.Sstart, file.Size, root.Sstart)
	} else {
		return o.stream_read(file.Sstart, file.Size)
	}
}

// Read MSAT
func (o *Ole) readMSAT() error {
	maxSID, sectorCount, err := o.maxRegularSID()
	if err != nil {
		return err
	}

	count := uint32(109)
	if o.header.Cfat < 109 {
		count = o.header.Cfat
	}

	var fatRead uint32

	for i := uint32(0); i < count && fatRead < o.header.Cfat; i++ {
		sid := o.header.Msat[i]
		if sid == FREESECT || sid == ENDOFCHAIN {
			break
		}
		if err := validateSID(sid, maxSID, "header MSAT FAT"); err != nil {
			return err
		}
		if sector, err := o.sector_read(sid); err == nil {
			sids := sector.AllValues(o.Lsector)
			o.SecID = append(o.SecID, sids...)
			fatRead++
		} else {
			return err
		}
	}

	// SAFETY: honor Cdif but cap it to file sector count.
	cdifLimit := o.header.Cdif
	if cdifLimit > sectorCount {
		cdifLimit = sectorCount
	}
	visitedDif := make(map[uint32]struct{})
	var difRead uint32
	for sid := o.header.Difstart; sid != ENDOFCHAIN && fatRead < o.header.Cfat; {
		if sid == FREESECT {
			break
		}
		if difRead >= cdifLimit {
			return fmt.Errorf("DIFAT chain exceeds Cdif limit: %d", cdifLimit)
		}
		if _, ok := visitedDif[sid]; ok {
			return fmt.Errorf("DIFAT chain cycle detected at sid: %d", sid)
		}
		visitedDif[sid] = struct{}{}
		if err := validateSID(sid, maxSID, "DIFAT sector"); err != nil {
			return err
		}
		if sector, err := o.sector_read(sid); err == nil {
			sids := sector.MsatValues(o.Lsector)

			for _, fatSID := range sids {
				if fatRead >= o.header.Cfat {
					break
				}
				if fatSID == FREESECT || fatSID == ENDOFCHAIN {
					continue
				}
				if err := validateSID(fatSID, maxSID, "DIFAT FAT"); err != nil {
					return err
				}
				if sector, err := o.sector_read(fatSID); err == nil {
					sids := sector.AllValues(o.Lsector)

					o.SecID = append(o.SecID, sids...)
					fatRead++
				} else {
					return err
				}
			}

			sid = sector.NextSid(o.Lsector)
			difRead++
		} else {
			return err
		}
	}
	if fatRead < o.header.Cfat {
		return fmt.Errorf("incomplete FAT chain: expected %d sectors, got %d", o.header.Cfat, fatRead)
	}

	for i := uint32(0); i < o.header.Csfat; i++ {
		sid := o.header.Sfatstart

		if sid != ENDOFCHAIN {
			if sector, err := o.sector_read(sid); err == nil {
				sids := sector.MsatValues(o.Lsector)

				o.SSecID = append(o.SSecID, sids...)

				sid = sector.NextSid(o.Lsector)
			} else {
				return err
			}
		}
	}
	return nil

}

func (o *Ole) maxRegularSID() (uint32, uint32, error) {
	current, err := o.reader.Seek(0, io.SeekCurrent)
	if err != nil {
		return 0, 0, err
	}
	end, err := o.reader.Seek(0, io.SeekEnd)
	if err != nil {
		return 0, 0, err
	}
	if _, err := o.reader.Seek(current, io.SeekStart); err != nil {
		return 0, 0, err
	}
	if end < int64(o.Lsector) {
		return 0, 0, fmt.Errorf("invalid OLE file size: %d", end)
	}

	sectorCount64 := uint64(end-int64(o.Lsector)) / uint64(o.Lsector)
	if sectorCount64 == 0 {
		return 0, 0, fmt.Errorf("empty OLE sector area")
	}
	if sectorCount64 > uint64(^uint32(0))+1 {
		return 0, 0, fmt.Errorf("too many sectors: %d", sectorCount64)
	}
	sectorCount := uint32(sectorCount64)
	return sectorCount - 1, sectorCount, nil
}

func validateSID(sid uint32, maxSID uint32, label string) error {
	if sid > maxSID {
		return fmt.Errorf("%s sid out of range: %d > %d", label, sid, maxSID)
	}
	return nil
}

func (o *Ole) stream_read(sid uint32, size uint32) *StreamReader {
	return &StreamReader{o.SecID, sid, o.reader, sid, 0, o.Lsector, int64(size), 0, sector_pos}
}

func (o *Ole) short_stream_read(sid uint32, size uint32, startSecId uint32) *StreamReader {
	ssatReader := &StreamReader{o.SecID, startSecId, o.reader, sid, 0, o.Lsector, int64(uint32(len(o.SSecID)) * o.Lssector), 0, sector_pos}
	return &StreamReader{o.SSecID, sid, ssatReader, sid, 0, o.Lssector, int64(size), 0, short_sector_pos}
}

func (o *Ole) sector_read(sid uint32) (Sector, error) {
	return o.sector_read_internal(sid, o.Lsector)
}

func (o *Ole) short_sector_read(sid uint32) (Sector, error) {
	return o.sector_read_internal(sid, o.Lssector)
}

func (o *Ole) sector_read_internal(sid, size uint32) (Sector, error) {
	pos := sector_pos(sid, size)
	if _, err := o.reader.Seek(int64(pos), 0); err == nil {
		var bts = make([]byte, size)
		o.reader.Read(bts)
		return Sector(bts), nil
	} else {
		return nil, err
	}
}

func sector_pos(sid uint32, size uint32) uint32 {
	return 512 + sid*size
}

func short_sector_pos(sid uint32, size uint32) uint32 {
	return sid * size
}

package excelize

import (
	"archive/zip"
	"bufio"
	"bytes"
	"encoding/binary"
	"encoding/xml"
	"errors"
	"fmt"
	"io"
	"io/fs"
	"math"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"sync"
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
)

// errZipWriter is a mock ZipWriter whose Create and Close can be configured to
// return errors.
type errZipWriter struct {
	createFunc func(string) (io.Writer, error)
	closeErr   error
}

func (m *errZipWriter) Create(name string) (io.Writer, error) {
	if m.createFunc != nil {
		return m.createFunc(name)
	}
	return &bytes.Buffer{}, nil
}

func (m *errZipWriter) AddFS(fs.FS) error { return nil }

func (m *errZipWriter) Close() error { return m.closeErr }

type errWriter struct{ err error }

func (e *errWriter) Write([]byte) (int, error) { return 0, e.err }

func BenchmarkWrite(b *testing.B) {
	const s = "This is test data"
	for b.Loop() {
		f := NewFile()
		for row := 1; row <= 10000; row++ {
			for col := 1; col <= 20; col++ {
				val, err := CoordinatesToCellName(col, row)
				if err != nil {
					b.Error(err)
				}
				if err := f.SetCellValue("Sheet1", val, s); err != nil {
					b.Error(err)
				}
			}
		}
		// Save spreadsheet by the given path.
		if err := f.SaveAs("test.xlsx"); err != nil {
			b.Error(err)
		}
	}
}

func TestWriteTo(t *testing.T) {
	// Test WriteToBuffer err
	{
		f, buf := File{Pkg: sync.Map{}}, bytes.Buffer{}
		f.SetZipWriter(func(w io.Writer) ZipWriter { return zip.NewWriter(w) })
		f.Pkg.Store("/d/", []byte("s"))
		_, err := f.WriteTo(bufio.NewWriter(&buf))
		assert.EqualError(t, err, "zip: write to directory")
		f.Pkg.Delete("/d/")
	}
	// Test file path overflow
	{
		f, buf := File{Pkg: sync.Map{}}, bytes.Buffer{}
		f.SetZipWriter(func(w io.Writer) ZipWriter { return zip.NewWriter(w) })
		const maxUint16 = 1<<16 - 1
		f.Pkg.Store(strings.Repeat("s", maxUint16+1), nil)
		_, err := f.WriteTo(bufio.NewWriter(&buf))
		assert.EqualError(t, err, "zip: FileHeader.Name too long")
	}
	// Test StreamsWriter err
	{
		f, buf := File{Pkg: sync.Map{}}, bytes.Buffer{}
		f.SetZipWriter(func(w io.Writer) ZipWriter { return zip.NewWriter(w) })
		f.Pkg.Store("s", nil)
		f.streams = make(map[string]*StreamWriter)
		file, _ := os.Open("123")
		f.streams["s"] = &StreamWriter{rawData: bufferedWriter{tmp: file}}
		_, err := f.WriteTo(bufio.NewWriter(&buf))
		assert.Nil(t, err)
	}
	// Test write with temporary file
	{
		f, buf := File{tempFiles: sync.Map{}}, bytes.Buffer{}
		f.SetZipWriter(func(w io.Writer) ZipWriter { return zip.NewWriter(w) })
		const maxUint16 = 1<<16 - 1
		f.tempFiles.Store("s", "")
		f.tempFiles.Store(strings.Repeat("s", maxUint16+1), "")
		_, err := f.WriteTo(bufio.NewWriter(&buf))
		assert.EqualError(t, err, "zip: FileHeader.Name too long")
	}
	// Test write with unsupported workbook file format
	{
		f, buf := File{Pkg: sync.Map{}}, bytes.Buffer{}
		f.SetZipWriter(func(w io.Writer) ZipWriter { return zip.NewWriter(w) })
		f.Pkg.Store("/d", []byte("s"))
		f.Path = "Book1.xls"
		_, err := f.WriteTo(bufio.NewWriter(&buf))
		assert.EqualError(t, err, ErrWorkbookFileFormat.Error())
	}
	// Test write with unsupported charset content types.
	{
		f, buf := NewFile(), bytes.Buffer{}
		f.ContentTypes, f.Path = nil, filepath.Join("test", "TestWriteTo.xlsx")
		f.Pkg.Store(defaultXMLPathContentTypes, MacintoshCyrillicCharset)
		_, err := f.WriteTo(bufio.NewWriter(&buf))
		assert.EqualError(t, err, "XML syntax error on line 1: invalid UTF-8")
	}
	// Test WriteToBuffer with ZipWriter Close error
	{
		f := NewFile()
		f.SetZipWriter(func(w io.Writer) ZipWriter {
			return &errZipWriter{closeErr: errors.New("close error")}
		})
		_, err := f.WriteTo(bufio.NewWriter(&bytes.Buffer{}))
		assert.EqualError(t, err, "close error")
	}
	// Test writeToZip with stream Create error
	{
		f := NewFile()
		f.streams = make(map[string]*StreamWriter)
		f.streams["s"] = &StreamWriter{rawData: bufferedWriter{}}
		f.SetZipWriter(func(w io.Writer) ZipWriter {
			return &errZipWriter{
				createFunc: func(name string) (io.Writer, error) {
					if name == "s" {
						return nil, errors.New("create stream error")
					}
					return &bytes.Buffer{}, nil
				},
			}
		})
		_, err := f.WriteTo(bufio.NewWriter(&bytes.Buffer{}))
		assert.EqualError(t, err, "create stream error")
	}
	// Test writeToZip with stream rawData.Reader() error
	{
		f := NewFile()
		f.streams = make(map[string]*StreamWriter)
		tmp, err := os.CreateTemp("", "excelize-test-*")
		assert.NoError(t, err)
		assert.NoError(t, tmp.Close())
		f.streams["s"] = &StreamWriter{rawData: bufferedWriter{tmp: tmp}}
		_, err = f.WriteTo(bufio.NewWriter(&bytes.Buffer{}))
		assert.Error(t, err)
	}
	// Test writeToZip with io.Copy error on stream
	{
		f := NewFile()
		f.streams = make(map[string]*StreamWriter)
		sw := &StreamWriter{}
		sw.rawData.buf.WriteString("test data")
		f.streams["s"] = sw
		f.SetZipWriter(func(w io.Writer) ZipWriter {
			return &errZipWriter{
				createFunc: func(name string) (io.Writer, error) {
					if name == "s" {
						return &errWriter{err: errors.New("copy error")}, nil
					}
					return &bytes.Buffer{}, nil
				},
			}
		})
		_, err := f.WriteTo(bufio.NewWriter(&bytes.Buffer{}))
		assert.EqualError(t, err, "copy error")
	}
}

func TestClose(t *testing.T) {
	f := NewFile()
	f.tempFiles.Store("/d/", "/d/")
	require.Error(t, f.Close())
}

func TestZip64(t *testing.T) {
	f := NewFile()
	_, err := f.NewSheet("Sheet2")
	assert.NoError(t, err)
	sw, err := f.NewStreamWriter("Sheet1")
	assert.NoError(t, err)
	for r := range 131 {
		rowData := make([]interface{}, 1000)
		for c := range 1000 {
			rowData[c] = strings.Repeat("c", TotalCellChars)
		}
		cell, err := CoordinatesToCellName(1, r+1)
		assert.NoError(t, err)
		assert.NoError(t, sw.SetRow(cell, rowData))
	}
	assert.NoError(t, sw.Flush())
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestZip64.xlsx")))
	assert.NoError(t, f.Close())

	// Test with filename length overflow
	f = NewFile()
	f.zip64Entries = append(f.zip64Entries, defaultXMLPathSharedStrings)
	buf := new(bytes.Buffer)
	buf.Write([]byte{0x50, 0x4b, 0x03, 0x04})
	buf.Write(make([]byte, 20))
	assert.NoError(t, f.writeZip64LFH(buf))

	// Test with file header less than the required 30 for the fixed header part
	f = NewFile()
	f.zip64Entries = append(f.zip64Entries, defaultXMLPathSharedStrings)
	buf.Reset()
	buf.Write([]byte{0x50, 0x4b, 0x03, 0x04})
	buf.Write(make([]byte, 22))
	assert.NoError(t, binary.Write(buf, binary.LittleEndian, uint16(10)))
	buf.Write(make([]byte, 2))
	buf.WriteString("test")
	assert.NoError(t, f.writeZip64LFH(buf))

	t.Run("for_save_zip64_with_in_memory_file_over_4GB", func(t *testing.T) {
		// Test save workbook in ZIP64 format with in memory file with size over 4GB.
		f := NewFile()
		f.Sheet.Delete("xl/worksheets/sheet1.xml")
		f.Pkg.Store("xl/worksheets/sheet1.xml", make([]byte, math.MaxUint32+1))
		_, err := f.WriteToBuffer()
		assert.NoError(t, err)
		assert.NoError(t, f.Close())
	})

	t.Run("for_save_zip64_with_in_temporary_file_over_4GB", func(t *testing.T) {
		// Test save workbook in ZIP64 format with temporary file with size over 4GB.
		if os.Getenv("GITHUB_ACTIONS") == "true" {
			t.Skip()
		}
		f := NewFile()
		f.Pkg.Delete("xl/worksheets/sheet1.xml")
		f.Sheet.Delete("xl/worksheets/sheet1.xml")
		tmp, err := os.CreateTemp(os.TempDir(), "excelize-")
		assert.NoError(t, err)
		assert.NoError(t, tmp.Truncate(math.MaxUint32+1))
		f.tempFiles.Store("xl/worksheets/sheet1.xml", tmp.Name())
		assert.NoError(t, tmp.Close())
		_, err = f.WriteToBuffer()
		assert.NoError(t, err)
		assert.NoError(t, f.Close())
	})
}

func TestRemoveTempFiles(t *testing.T) {
	tmp, err := os.CreateTemp("", "excelize-*")
	if err != nil {
		t.Fatal(err)
	}
	tmpName := tmp.Name()
	assert.NoError(t, tmp.Close())
	f := NewFile()
	// Fill the tempFiles map with non-existing files
	for i := range 1000 {
		f.tempFiles.Store(strconv.Itoa(i), "/hopefully not existing")
	}
	f.tempFiles.Store("existing", tmpName)

	require.Error(t, f.Close())
	if _, err := os.Stat(tmpName); err == nil {
		t.Errorf("temp file %q still exist", tmpName)
		assert.NoError(t, os.Remove(tmpName))
	}
}

func TestWriteToBufferProgress(t *testing.T) {
	type report struct{ processed, total int64 }
	var reports []report
	progress := func(processed, total int64) {
		reports = append(reports, report{processed, total})
	}
	check := func(t *testing.T) {
		require.NotEmpty(t, reports)
		assert.Greater(t, reports[0].total, int64(0)) // estimated total
		for i := 1; i < len(reports); i++ {
			assert.GreaterOrEqual(t, reports[i].processed, reports[i-1].processed)
		}
		last := reports[len(reports)-1]
		assert.Equal(t, last.total, last.processed)
		assert.Greater(t, last.total, int64(0))
	}

	// Test save progress with the normal writer
	f := NewFile(Options{SaveProgress: progress})
	for row := 1; row <= 100; row++ {
		for col := 1; col <= 10; col++ {
			cell, err := CoordinatesToCellName(col, row)
			require.NoError(t, err)
			require.NoError(t, f.SetCellValue("Sheet1", cell, row*col))
		}
	}
	buf, err := f.WriteToBuffer()
	require.NoError(t, err)
	check(t)
	_, err = zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	assert.NoError(t, err)

	// Test save progress is reset on the second save
	reports = reports[:0]
	_, err = f.WriteToBuffer()
	require.NoError(t, err)
	check(t)
	require.NoError(t, f.Close())

	// Test save progress with the stream writer
	reports = reports[:0]
	f = NewFile(Options{SaveProgress: progress})
	sw, err := f.NewStreamWriter("Sheet1")
	require.NoError(t, err)
	for row := 1; row <= 100; row++ {
		cell, err := CoordinatesToCellName(1, row)
		require.NoError(t, err)
		require.NoError(t, sw.SetRow(cell, []interface{}{row}))
	}
	require.NoError(t, sw.Flush())
	_, err = f.WriteToBuffer()
	require.NoError(t, err)
	check(t)
	require.NoError(t, f.Close())

	// Test save progress with password protection
	reports = reports[:0]
	f = NewFile(Options{Password: "password", SaveProgress: progress})
	require.NoError(t, f.SetCellValue("Sheet1", "A1", "Hello"))
	_, err = f.WriteToBuffer()
	require.NoError(t, err)
	check(t)
	require.NoError(t, f.Close())
}

func TestProgressWriter(t *testing.T) {
	var calls int
	var processed, total int64
	pw := &progressWriter{
		callback:  func(p, t int64) { calls++; processed, total = p, t },
		total:     210,
		threshold: 64,
	}
	// Report nothing when the progress does not exceed the threshold
	_, err := pw.Write(make([]byte, 50))
	assert.NoError(t, err)
	assert.Zero(t, calls)
	// Report the progress clamped to the total when overflowing
	_, err = pw.Write(make([]byte, 200))
	assert.NoError(t, err)
	assert.Equal(t, 1, calls)
	assert.Equal(t, int64(210), processed)
	assert.Equal(t, int64(210), total)
	// Report nothing when the progress does not change
	pw.report(210)
	assert.Equal(t, 1, calls)
}

func TestSerializeProgressWriter(t *testing.T) {
	var reports []int64
	f := NewFile(Options{SaveProgress: func(processed, total int64) {
		assert.Equal(t, processed, total) // clamped to processed without an estimate
		reports = append(reports, processed)
	}})
	defer func() { assert.NoError(t, f.Close()) }()

	// No wrapper when the file is not being saved
	buf := bytes.Buffer{}
	assert.Equal(t, &buf, f.wrapSaveProgress(&buf))

	f.inSave = true
	w := f.wrapSaveProgress(&buf)
	// Report nothing when the progress does not exceed the threshold
	_, err := w.Write(make([]byte, 1000))
	assert.NoError(t, err)
	assert.Empty(t, reports)
	// Report the progress with pending bytes included
	_, err = w.Write(make([]byte, 100<<10))
	assert.NoError(t, err)
	assert.Equal(t, []int64{100<<10 + 1000}, reports)
	// Commit the pending bytes on saveFileList
	f.saveFileList("docProps/core.xml", buf.Bytes())
	assert.Equal(t, []int64{100<<10 + 1000, int64(buf.Len()) + int64(len(xml.Header))}, reports)
	assert.Zero(t, f.savePending)
}

func TestWriteToBufferEstimatedTotal(t *testing.T) {
	var reports [][2]int64
	f := NewFile(Options{SaveProgress: func(processed, total int64) {
		reports = append(reports, [2]int64{processed, total})
	}})
	for row := 1; row <= 2000; row++ {
		for col := 1; col <= 10; col++ {
			cell, err := CoordinatesToCellName(col, row)
			require.NoError(t, err)
			require.NoError(t, f.SetCellValue("Sheet1", cell, fmt.Sprintf("order-%d-%d", row, col)))
		}
	}
	_, err := f.WriteToBuffer()
	require.NoError(t, err)
	require.NoError(t, f.Close())
	require.NotEmpty(t, reports)
	estimated, final := reports[0][1], reports[len(reports)-1]
	assert.Positive(t, estimated)
	assert.Equal(t, final[1], final[0])
	// The estimated total is close to the exact total
	ratio := float64(estimated) / float64(final[1])
	t.Logf("estimated total %d, exact total %d, ratio %.3f", estimated, final[1], ratio)
	assert.Greater(t, ratio, 0.7)
	assert.Less(t, ratio, 1.4)
}

func TestEstimateSaveTotal(t *testing.T) {
	f := NewFile()
	base := f.estimateSaveTotal()
	assert.Greater(t, base, int64(0))
	// The estimate grows with the amount of data to be serialized
	for row := 1; row <= 500; row++ {
		for col := 1; col <= 10; col++ {
			cell, err := CoordinatesToCellName(col, row)
			require.NoError(t, err)
			require.NoError(t, f.SetCellValue("Sheet1", cell, "estimate me"))
		}
	}
	grown := f.estimateSaveTotal()
	assert.Greater(t, grown, base)
	// The row slice of the worksheet is restored after sampling
	_, err := f.WriteToBuffer()
	require.NoError(t, err)
	rows, err := f.GetRows("Sheet1")
	require.NoError(t, err)
	assert.Len(t, rows, 500)
	require.NoError(t, f.Close())

	// Estimate the size of a stream written worksheet exactly
	f = NewFile()
	sw, err := f.NewStreamWriter("Sheet1")
	require.NoError(t, err)
	require.NoError(t, sw.SetRow("A1", []interface{}{"stream"}))
	require.NoError(t, sw.Flush())
	streamEst := f.estimateSaveTotal()
	assert.Greater(t, streamEst, int64(0))
	require.NoError(t, f.Close())
}

package xls

import (
	"fmt"
	"math"
	"math/big"
	"strconv"
	"strings"
	"time"

	"github.com/morpon/extrame-xls/yymmdd"
)

// content type
type contentHandler interface {
	String(*WorkBook) []string
	FirstCol() uint16
	LastCol() uint16
}

type Col struct {
	RowB      uint16
	FirstColB uint16
}

type Coler interface {
	Row() uint16
}

func (c *Col) Row() uint16 {
	return c.RowB
}

func (c *Col) FirstCol() uint16 {
	return c.FirstColB
}

func (c *Col) LastCol() uint16 {
	return c.FirstColB
}

func (c *Col) String(wb *WorkBook) []string {
	return []string{"default"}
}

type XfRk struct {
	Index uint16
	Rk    RK
}

func (xf *XfRk) String(wb *WorkBook) string {
	idx := int(xf.Index)
	if len(wb.Xfs) > idx {
		fNo := wb.Xfs[idx].formatNo()
		if fNo >= 164 { // user defined format
			if formatter := wb.Formats[fNo]; formatter != nil {
				formatterLower := strings.ToLower(formatter.str)
				if formatterLower == "general" ||
					strings.Contains(formatter.str, "#") ||
					strings.Contains(formatter.str, ".00") ||
					strings.Contains(formatterLower, "m/y") ||
					strings.Contains(formatterLower, "d/y") ||
					strings.Contains(formatterLower, "m.y") ||
					strings.Contains(formatterLower, "d.y") ||
					strings.Contains(formatterLower, "h:") ||
					strings.Contains(formatterLower, "д.г") {
					//If format contains # or .00 then this is a number
					return xf.Rk.String()
				} else {
					i, f, isFloat := xf.Rk.number()
					if !isFloat {
						f = float64(i)
					}
					t := timeFromExcelTime(f, wb.dateMode == 1)
					return yymmdd.Format(t, formatter.str)
				}
			}
			// see http://www.openoffice.org/sc/excelfileformat.pdf Page #174
		} else if 14 <= fNo && fNo <= 17 || fNo == 22 || 27 <= fNo && fNo <= 36 || 50 <= fNo && fNo <= 58 { // jp. date format
			i, f, isFloat := xf.Rk.number()
			if !isFloat {
				f = float64(i)
			}
			t := timeFromExcelTime(f, wb.dateMode == 1)
			return t.Format(time.RFC3339) //TODO it should be international
		}
	}
	return xf.Rk.String()
}

type RK uint32

func (rk RK) number() (intNum int64, floatNum float64, isFloat bool) {
	multiplied := rk & 1
	isInt := rk & 2
	val := int32(rk) >> 2
	if isInt == 0 {
		isFloat = true
		floatNum = math.Float64frombits(uint64(val) << 34)
		if multiplied != 0 {
			floatNum = floatNum / 100
		}
		return
	}
	if multiplied != 0 {
		isFloat = true
		floatNum = float64(val) / 100
		return
	}
	return int64(val), 0, false
}

func (rk RK) String() string {
	i, f, isFloat := rk.number()
	if isFloat {
		return strconv.FormatFloat(f, 'f', -1, 64)
	}
	return strconv.FormatInt(i, 10)
}

var ErrIsInt = fmt.Errorf("is int")

func (rk RK) Float() (float64, error) {
	_, f, isFloat := rk.number()
	if !isFloat {
		return 0, ErrIsInt
	}
	return f, nil
}

type MulrkCol struct {
	Col
	Xfrks    []XfRk
	LastColB uint16
}

func (c *MulrkCol) LastCol() uint16 {
	return c.LastColB
}

func (c *MulrkCol) String(wb *WorkBook) []string {
	var res = make([]string, len(c.Xfrks))
	for i := 0; i < len(c.Xfrks); i++ {
		xfrk := c.Xfrks[i]
		res[i] = xfrk.String(wb)
	}
	return res
}

type MulBlankCol struct {
	Col
	Xfs      []uint16
	LastColB uint16
}

func (c *MulBlankCol) LastCol() uint16 {
	return c.LastColB
}

func (c *MulBlankCol) String(wb *WorkBook) []string {
	return make([]string, len(c.Xfs))
}

type NumberCol struct {
	Col
	Index uint16
	Float float64
}

func (c *NumberCol) String(wb *WorkBook) []string {
	fNo := wb.Xfs[c.Index].formatNo()
	if fNo != 0 && wb.Formats[fNo] != nil {
		numFmt := wb.Formats[fNo].str
		if IsDateTimeFormat(numFmt) {
			t := timeFromExcelTime(c.Float, wb.dateMode == 1)
			// 根据数字格式判断具体的日期时间类型
			if IsDateOnlyFormat(numFmt) {
				// 纯日期格式
				return []string{t.Format("2006-01-02")}
			} else if IsTimeOnlyFormat(numFmt) {
				// 纯时间格式
				return []string{t.Format("15:04:05")}
			} else {
				// 日期时间格式
				return []string{t.Format("2006-01-02 15:04:05")}
			}
		}
	}

	floatValue := c.Float
	// 检查是否为整数
	intValue := float64(int64(floatValue))
	if floatValue == intValue || (floatValue <= 9007199254740992 && floatValue >= -9007199254740992) {
		return []string{strconv.FormatFloat(floatValue, 'f', -1, 64)}
	} else {
		// 使用big.Float处理超大精度数字
		bf := new(big.Float).SetFloat64(floatValue).SetPrec(256)
		return []string{bf.Text('f', -1)}
	}
}

type FormulaStringCol struct {
	Col
	RenderedValue string
}

func (c *FormulaStringCol) String(wb *WorkBook) []string {
	return []string{c.RenderedValue}
}

//str, err = wb.get_string(buf_item, size)
//wb.sst[offset_pre] = wb.sst[offset_pre] + str

type FormulaCol struct {
	Header struct {
		Col
		IndexXf uint16
		Result  [8]byte
		Flags   uint16
		_       uint32
	}
	Bts []byte
}

func (c *FormulaCol) String(wb *WorkBook) []string {
	return []string{"FormulaCol"}
}

type RkCol struct {
	Col
	Xfrk XfRk
}

func (c *RkCol) String(wb *WorkBook) []string {
	return []string{c.Xfrk.String(wb)}
}

type LabelsstCol struct {
	Col
	Xf  uint16
	Sst uint32
}

func (c *LabelsstCol) String(wb *WorkBook) []string {
	return []string{wb.sst[int(c.Sst)]}
}

type labelCol struct {
	BlankCol
	Str string
}

func (c *labelCol) String(wb *WorkBook) []string {
	return []string{c.Str}
}

type BlankCol struct {
	Col
	Xf uint16
}

func (c *BlankCol) String(wb *WorkBook) []string {
	return []string{""}
}

// isDateTimeFormat 判断数字格式是否为日期时间格式
func IsDateTimeFormat(numFmt string) bool {
	if numFmt == "" {
		return false
	}

	// 常见的日期时间格式模式
	dateTimePatterns := []string{
		"yyyy", "yy", "mm", "dd", "hh", "ss",
		"m/d", "d/m", "yyyy-mm-dd", "dd-mm-yyyy", "mm-dd-yyyy",
		"h:mm", "hh:mm:ss", "mm:ss", "h:mm:ss",
		"yyyy/mm/dd", "dd/mm/yyyy", "mm/dd/yyyy",
		"general date", "long date", "medium date", "short date",
		"long time", "medium time", "short time",
	}

	numFmtLower := strings.ToLower(numFmt)
	for _, pattern := range dateTimePatterns {
		if strings.Contains(numFmtLower, pattern) {
			return true
		}
	}

	return false
}

// IsDateOnlyFormat 判断是否为纯日期格式（不包含时间）
func IsDateOnlyFormat(numFmt string) bool {
	if numFmt == "" {
		return false
	}

	numFmtLower := strings.ToLower(numFmt)

	// 明确的时间模式 - 只判断小时和秒，避免mm歧义
	hasTimePattern := strings.Contains(numFmtLower, "h") || strings.Contains(numFmtLower, "hh") || strings.Contains(numFmtLower, "ss")

	// 明确的日期模式 - 直接判断年、月、日标识符
	hasYearPattern := strings.Contains(numFmtLower, "yyyy") || strings.Contains(numFmtLower, "yy")
	hasDayPattern := strings.Contains(numFmtLower, "dd") || strings.Contains(numFmtLower, "d")
	// mm既可能是月份也可能是分钟，但如果没有时间标识符，优先认为是月份
	hasMonthPattern := (strings.Contains(numFmtLower, "mm") || strings.Contains(numFmtLower, "m")) && !hasTimePattern

	// 必须有明确的日期成分，且没有时间成分
	hasDatePattern := (hasYearPattern || hasDayPattern || hasMonthPattern) &&
		(hasYearPattern || (hasMonthPattern && hasDayPattern))

	return hasDatePattern && !hasTimePattern
}

// IsTimeOnlyFormat 判断是否为纯时间格式（不包含日期）
func IsTimeOnlyFormat(numFmt string) bool {
	if numFmt == "" {
		return false
	}

	numFmtLower := strings.ToLower(numFmt)

	// 明确的时间模式 - 只判断小时和秒，避免mm歧义
	hasTimePattern := strings.Contains(numFmtLower, "h") ||
		strings.Contains(numFmtLower, "hh") ||
		strings.Contains(numFmtLower, "ss")

	// 明确的日期模式 - 直接判断年、月、日标识符
	hasYearPattern := strings.Contains(numFmtLower, "yyyy") || strings.Contains(numFmtLower, "yy")
	hasDayPattern := strings.Contains(numFmtLower, "dd") || strings.Contains(numFmtLower, "d")
	// mm既可能是月份也可能是分钟，但如果没有时间标识符，优先认为是月份
	hasMonthPattern := (strings.Contains(numFmtLower, "mm") || strings.Contains(numFmtLower, "m")) && !hasTimePattern

	hasDatePattern := hasYearPattern || hasDayPattern || hasMonthPattern

	return hasTimePattern && !hasDatePattern
}

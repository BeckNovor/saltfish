package main

import (
	"bufio"
	"encoding/base64"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/go-mail/mail"
	"github.com/google/uuid"
	"io"
	"io/ioutil"
	"log"
	"math"
	"math/rand"
	"net/http"
	"os"
	"regexp"
	"strconv"
	"strings"
	"time"
)

// ExchangeRate 定义汇率响应结构体
type ExchangeRate struct {
	Result   string             `json:"result"`
	BaseCode string             `json:"base_code"`
	Rates    map[string]float64 `json:"conversion_rates"`
}

// FetchExchangeRate 函数用于从指定 URL 抓取 HTML 代码
func FetchExchangeRate() float64 {
	url := "https://www.nbs.rs/static/nbs_site/gen/english/30/kurs/Indikativni_Kurs_20.html"
	// 发起 HTTP GET 请求
	resp, _ := http.Get(url)

	// 确保在函数结束时关闭响应体
	defer resp.Body.Close()

	// 读取响应体内容
	body, _ := ioutil.ReadAll(resp.Body)

	// 将响应体内容转换为字符串
	html := string(body)
	re := regexp.MustCompile(`<th class="kurs_e">([\d.]+)</th>`)
	html_txt := re.FindStringSubmatch(html)
	exchange_rate, _ := strconv.ParseFloat(html_txt[1], 64)
	return exchange_rate
}

// numberToExcelColumn 函数将输入的数字转换为 Excel 表头对应的字母
func numberToExcelColumn(num int) string {
	if num <= 0 {
		return ""
	}
	var result string
	for num > 0 {
		// 计算当前位对应的字母偏移量，这里减去 1 是为了从 0 开始计数
		remainder := (num - 1) % 26
		// 将偏移量转换为对应的字母并添加到结果字符串的前面
		result = string('A'+remainder) + result
		// 更新 num 为去掉当前位后的数字
		num = (num - 1) / 26
	}
	return result
}

// 需求 1：检查单元格 AD1 是否是客户单号，是的话删除 AD 列
func deleteADColumn(column int, mf_file *excelize.File) {
	ADColumn := numberToExcelColumn(column)
	cellAD1 := mf_file.GetCellValue("Sheet1", ADColumn+"1")
	if cellAD1 == "客户单号" {
		mf_file.RemoveCol("Sheet1", ADColumn)
	}
}

// 生成随机字母和数字的字符串
func randomString() string {
	id := uuid.New()
	return id.String()
}

// GenerateRandomString 生成指定长度的随机字母和数字字符串
func GenerateRandomString(length int) string {
	// 计算需要的字节数
	numBytes := (length*3 + 3) / 4
	// 创建字节切片来存储随机字节
	bytes := make([]byte, numBytes)
	// 从操作系统的加密随机数生成器中读取随机字节
	_, err := rand.Read(bytes)
	if err != nil {
		return ""
	}
	// 对随机字节进行 Base64 编码
	encoded := base64.URLEncoding.EncodeToString(bytes)
	// 截取指定长度的字符串
	return encoded[:length]
}

func excelSerialToTime(serialStr string) (time.Time, error) {
	serial, err := strconv.ParseFloat(serialStr, 64)
	if err != nil {
		return time.Time{}, err
	}
	days := int(serial)
	// 修正 Excel 对 1900 年闰年的错误处理
	if days > 59 {
		days--
	}
	// 计算一天内的时间偏移量
	secs := int((serial - float64(days)) * 24 * 60 * 60)
	// 从 1899 年 12 月 30 日开始计算
	return time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC).
		AddDate(0, 0, days).
		Add(time.Duration(secs) * time.Second), nil
}

// 需求 2：检查  列最后一个单元格是否为空，为空则删除最后一行
func processLastRow(ioss_column int, mf_rows [][]string, mf_file *excelize.File) {
	lastRow := len(mf_rows)
	if lastRow > 0 {
		cellTLast := mf_file.GetCellValue("Sheet1", numberToExcelColumn(ioss_column)+strconv.Itoa(lastRow))
		if len(cellTLast) < 1 {
			mf_file.RemoveRow("Sheet1", lastRow-1)
			lastRow--
		}
	}
}

// 需求 3：处理 M 列和 N 列
func processMNColumns(buyerName int, buyerAdd int, tracking_number int, mf_file *excelize.File, mf_rows [][]string) {
	// 用于存储单号对应的买家信息
	orderInfo := make(map[string][2]string)

	for i, mf_row := range mf_rows {
		if i == 0 {
			continue
		}
		orderNumber := mf_row[tracking_number-1] // tracking number单号
		mValue := mf_row[buyerName-1]
		nValue := mf_row[buyerAdd-1]
		// 截取前 50 字符
		if len(mValue) > 50 {
			mValue = mValue[:50]
		}
		if len(nValue) > 50 {
			nValue = nValue[:50]
		}
		if _, exists := orderInfo[orderNumber]; !exists {
			// 如果该单号还没有对应的买家信息，生成新的名字和地址
			if mValue == "" { // M 列
				mValue = randomString()
			}
			if nValue == "" { // N 列
				nValue = randomString()

				// 地址里没有阿拉伯数字就在末尾加一个 0
				if !strings.ContainsAny(nValue, "0123456789") {
					nValue = nValue + "0"
				}
			} else {
				if !strings.ContainsAny(nValue, "0123456789") {
					nValue = nValue + "0"
				} else {
					nValue = nValue
				}
			}

			// 存储该单号对应的买家信息
			orderInfo[orderNumber] = [2]string{mValue, nValue}
		}
		// 设置该行的买家名字和地址
		mf_file.SetCellValue("Sheet1", numberToExcelColumn(buyerName)+strconv.Itoa(i+1), orderInfo[orderNumber][0])
		mf_file.SetCellValue("Sheet1", numberToExcelColumn(buyerAdd)+strconv.Itoa(i+1), orderInfo[orderNumber][1])
	}
}

// 需求 4：处理金额限制
func processAmountLimit(buyerNameColumn int, buyerAddColumn int, priceColumn int, trackingNumber int, mf_file *excelize.File, mf_rows [][]string) {
	// 用于存储买家信息和总金额
	buyerInfo := make(map[string]struct {
		mValue       string
		nValue       string
		amount       float64
		rows         []int
		orderNumbers map[string]bool // 记录涉及的 AC 列单号
	})

	for i, mf_row := range mf_rows {
		if i == 0 {
			continue
		}
		mValue := mf_row[buyerNameColumn-1]                        // M 列买家名字
		nValue := mf_row[buyerAddColumn-1]                         // N 列买家地址
		amount, _ := strconv.ParseFloat(mf_row[priceColumn-1], 64) // U 列 TotalPrice
		orderNumber := mf_row[trackingNumber-1]                    // AC 列单号

		buyerKey := mValue + nValue
		if info, exists := buyerInfo[buyerKey]; exists {
			// 如果买家信息已存在，累加金额并记录行号和单号
			info.amount += amount
			info.rows = append(info.rows, i+1)
			info.orderNumbers[orderNumber] = true
			buyerInfo[buyerKey] = info
		} else {
			// 如果买家信息不存在，初始化买家信息
			orderNumbers := make(map[string]bool)
			orderNumbers[orderNumber] = true
			buyerInfo[buyerKey] = struct {
				mValue       string
				nValue       string
				amount       float64
				rows         []int
				orderNumbers map[string]bool
			}{
				mValue:       mValue,
				nValue:       nValue,
				amount:       amount,
				rows:         []int{i + 1},
				orderNumbers: orderNumbers,
			}
		}
	}
	for _, info := range buyerInfo {
		if info.amount > 150 {
			// 为每个单号生成唯一的随机字符串
			orderNumberRandomMap := make(map[string]string)
			for orderNumber := range info.orderNumbers {
				randomChar := GenerateRandomString(5)
				orderNumberRandomMap[orderNumber] = randomChar
			}
			//fmt.Println(orderNumberRandomMap)

			// 直接根据记录的行号更新买家信息
			for _, rowNum := range info.rows {
				orderNumber := mf_rows[rowNum-1][trackingNumber-1] // 获取该行对应的 AC 列单号
				randomChar := orderNumberRandomMap[orderNumber]
				newMValue := info.mValue + randomChar
				newNValue := info.nValue + randomChar
				mf_file.SetCellValue("Sheet1", numberToExcelColumn(buyerNameColumn)+strconv.Itoa(rowNum), newMValue)
				mf_file.SetCellValue("Sheet1", numberToExcelColumn(buyerAddColumn)+strconv.Itoa(rowNum), newNValue)
			}
		}
	}
}

func BEGprocessAmountLimit(mf_file *excelize.File, mf_rows [][]string) {
	// 用于存储每个单号的总金额和涉及的行号
	orderInfo := make(map[string]struct {
		amount float64
		rows   []int
	})

	// 遍历每一行数据
	for i, mf_row := range mf_rows {
		if i == 0 {
			continue
		}
		orderNumber := mf_row[1]                        // AC 列单号
		amount, _ := strconv.ParseFloat(mf_row[16], 64) // Q 列 TotalPrice

		if info, exists := orderInfo[orderNumber]; exists {
			// 如果单号信息已存在，累加金额并记录行号
			info.amount += amount
			info.rows = append(info.rows, i+1)
			orderInfo[orderNumber] = info
		} else {
			// 如果单号信息不存在，初始化单号信息
			orderInfo[orderNumber] = struct {
				amount float64
				rows   []int
			}{
				amount: amount,
				rows:   []int{i + 1},
			}
		}
	}

	// 计算金额限制
	limit := 49.5 * FetchExchangeRate()
	fmt.Println("封顶金额限制", limit)

	// 处理超过限制的单号
	for _, info := range orderInfo {
		if info.amount > limit {
			fmt.Println(info, "超过49.5欧限制了")
			// 计算缩小比例
			ratio := limit / info.amount
			// 遍历该单号涉及的所有行，缩小金额
			for _, rowNum := range info.rows {
				oldAmount, _ := strconv.ParseFloat(mf_rows[rowNum-1][16], 64)
				newAmount := math.Floor(oldAmount * ratio)
				// 更新 Excel 文件中 Q 列的金额
				mf_file.SetCellValue("Sheet1", "Q"+strconv.Itoa(rowNum), newAmount)
				// 更新内存中的数据
				mf_rows[rowNum-1][16] = strconv.FormatFloat(newAmount, 'f', 0, 64)
			}
		}
	}
}

// 需求 6：剔除 AC 列中的 %, 清空ConsigneeID和UNLOcode列,
func processACColumn(arrivalPort string, mf_file *excelize.File, mf_rows [][]string) {
	for i, mf_row := range mf_rows {
		if i == 0 {
			continue
		}
		//非LGG的清空ConsigneeID和UNLOcode列
		if arrivalPort != "LGG" {
			RcellValue := mf_file.GetCellValue("Sheet1", "R1")
			VcellValue := mf_file.GetCellValue("Sheet1", "V1")
			if RcellValue == "ConsigneeID" || VcellValue == "UNLOcode" {
				mf_file.SetCellValue("Sheet1", "R"+strconv.Itoa(i+1), "")
				mf_file.SetCellValue("Sheet1", "V"+strconv.Itoa(i+1), "")
			}
		}
		//剔除 AC 列中的 % & m列收件人地址
		acValue := mf_row[28]                                // AC列物流跟踪号
		total_price, _ := strconv.ParseFloat(mf_row[20], 64) // U列total_price
		ConsigneeName := mf_row[12]                          // N列收件人地址
		if strings.Contains(acValue, "%") || strings.Contains(ConsigneeName, " kft") {
			acValue = strings.ReplaceAll(acValue, "%", "")
			ConsigneeName = strings.ReplaceAll(ConsigneeName, " kft", "")
			mf_file.SetCellValue("Sheet1", "AC"+strconv.Itoa(i+1), acValue)
			mf_file.SetCellValue("Sheet1", "M"+strconv.Itoa(i+1), ConsigneeName)
		}

		ConsignorCity := mf_row[9] // J列ConsignorCity
		if len(ConsignorCity) < 1 {
			mf_file.SetCellValue("Sheet1", "J"+strconv.Itoa(i+1), "zhaoqing") // J列ConsignorCity
			mf_file.SetCellValue("Sheet1", "K"+strconv.Itoa(i+1), "526200")   //K列ConsignorPostcode
			mf_file.SetCellValue("Sheet1", "F"+strconv.Itoa(i+1), i)          //F列序号重新编号
			mf_file.SetCellValue("Sheet1", "U"+strconv.Itoa(i+1), total_price)

		}
	}
}

// 需求 6：填补发件人城市&邮编
func processOTPACColumn(arrivalPort string, mf_file *excelize.File, mf_rows [][]string) {
	for i, mf_row := range mf_rows {
		if i == 0 {
			continue
		}
		ConsignorCity := mf_row[10] // k列ConsignorCity
		if len(ConsignorCity) < 1 {
			mf_file.SetCellValue("Sheet1", "J"+strconv.Itoa(i+1), "526200")   // J列ConsignorCity
			mf_file.SetCellValue("Sheet1", "K"+strconv.Itoa(i+1), "zhaoqing") //K列ConsignorPostcode

		}
	}
}

// 需求 5：处理 LGG 提单列表.xlsx，重量分摊
func processLGGFile(NetMassColumn int, line int, mf_rows [][]string, awblist_rows [][]string, mf_file *excelize.File) {

	// 计算净重和
	netWeightSum := 0.0
	expansionCoefficient := 1.0
	for i, mf_row := range mf_rows {
		if i == 0 {
			continue
		}
		mf_NetMassKg := mf_row[NetMassColumn-1] //W列NetMassKg
		netweight, _ := strconv.ParseFloat(mf_NetMassKg, 64)
		if netweight < 0.01 {
			netweight = 0.01
		}
		mf_file.SetCellValue("Sheet1", numberToExcelColumn(NetMassColumn)+strconv.Itoa(i+1), netweight)
		netWeightSum += netweight
	}

	// 计算膨胀系数并填充 X 列
	billableWeight, _ := strconv.ParseFloat(awblist_rows[line][11], 64) //awblist_row[11] →提单计费重列
	expansionCoefficient = (billableWeight - 1) / netWeightSum

	for i, mf_row := range mf_rows {
		if i == 0 {
			continue
		}
		mf_NetMassKg := mf_row[NetMassColumn-1] //W列NetMassKg
		netweight, _ := strconv.ParseFloat(mf_NetMassKg, 64)
		xValue, _ := strconv.ParseFloat(fmt.Sprintf("%.3f", expansionCoefficient*netweight), 64)
		if xValue < 0.01 {
			xValue = 0.01
		}
		mf_file.SetCellValue("Sheet1", "X"+strconv.Itoa(i+1), xValue)

	}
	// awblist_file.Save()

}

// 需求 5：处理 OTP 提单列表.xlsx，重量分摊
func processOTPFile(NetMassColumn int, line int, mf_rows [][]string, awblist_rows [][]string, mf_file *excelize.File) {

	// 计算净重和
	netWeightSum := 0.0
	expansionCoefficient := 1.0
	for i, mf_row := range mf_rows {
		if i == 0 {
			continue
		}
		mf_NetMassKg := mf_row[NetMassColumn-1] //W列NetMassKg
		netweight, _ := strconv.ParseFloat(mf_NetMassKg, 64)
		netWeightSum += netweight
	}

	// 计算膨胀系数并填充 X 列
	billableWeight, _ := strconv.ParseFloat(awblist_rows[line][11], 64) //awblist_row[11] →提单计费重列
	expansionCoefficient = (billableWeight - 1) / netWeightSum

	for i, mf_row := range mf_rows {
		if i == 0 {
			continue
		}
		mf_NetMassKg := mf_row[NetMassColumn-1] //W列NetMassKg
		netweight, _ := strconv.ParseFloat(mf_NetMassKg, 64)
		xValue, _ := strconv.ParseFloat(fmt.Sprintf("%.3f", expansionCoefficient*netweight), 64)
		if xValue < 0.01 {
			xValue = 0.01
		}
		mf_file.SetCellValue("Sheet1", numberToExcelColumn(NetMassColumn)+strconv.Itoa(i+1), xValue)
		// 更新 mf_rows 中的值为膨胀后的新值
		mf_rows[i][NetMassColumn-1] = strconv.FormatFloat(xValue, 'f', 3, 64)
	}
	// 复制 Sheet1 为 package 表
	fromIndex := mf_file.GetSheetIndex("Sheet1")
	toIndex := mf_file.NewSheet("package")
	mf_file.CopySheet(fromIndex, toIndex)

	// 找到 WEIGHT 18 04 000 000 和 VALUE 14 14 000 000 列的索引
	weightColumnIndex := 4
	valueColumnIndex := 22

	// 按 A 列包裹号合计重量和金额
	packageWeightMap := make(map[string]float64)
	packageValueMap := make(map[string]float64)
	for i, mf_row := range mf_rows {
		if i == 0 {
			continue
		}
		packageNumber := mf_row[0] // A 列包裹号
		weightStr := mf_row[weightColumnIndex]
		valueStr := mf_row[valueColumnIndex]

		weight, _ := strconv.ParseFloat(weightStr, 64)
		value, _ := strconv.ParseFloat(valueStr, 64)

		packageWeightMap[packageNumber] += weight
		packageValueMap[packageNumber] += value
	}

	// 将合计结果写入 package 表
	for i, mf_row := range mf_rows {
		if i == 0 {
			continue
		}
		packageNumber := mf_row[0] // A 列包裹号
		mf_file.SetCellValue("package", numberToExcelColumn(weightColumnIndex+1)+strconv.Itoa(i+1), packageWeightMap[packageNumber])
		mf_file.SetCellValue("package", numberToExcelColumn(valueColumnIndex+1)+strconv.Itoa(i+1), packageValueMap[packageNumber])
	}
}

// 需求 7：处理 HSH 预报表格.xlsx
func processHSHFile(layout string, awb string, line int, awblist_rows [][]string) {
	file_hsh, err := excelize.OpenFile("HSH预报表格.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
	file_hsh.SetCellValue("Sheet1", "I1", "ETD")

	ETD, _ := excelSerialToTime(awblist_rows[line][5])
	ETA, _ := excelSerialToTime(awblist_rows[line][6])
	Boxes, _ := strconv.ParseFloat(awblist_rows[line][7], 64)
	Parcels, _ := strconv.ParseFloat(awblist_rows[line][10], 64)
	Chargeable_Weight, _ := strconv.ParseFloat(awblist_rows[line][11], 64)

	values := []interface{}{
		"Airfreight", "YDH", "B2C", awblist_rows[line][0], awblist_rows[line][0], "", awblist_rows[line][1], awblist_rows[line][4], ETD.Format(layout), ETA.Format(layout), Boxes, Parcels, Chargeable_Weight, awblist_rows[line][18], "HSH", "BECUGHE000048", "BECUGHE000048",
	}
	for j, value := range values {
		file_hsh.SetCellValue("Sheet1", string(rune('A'+j))+"2", value)
	}
	newHshFileName := awb + "_" + "HSH预报表格.xlsx"
	file_hsh.SaveAs(newHshFileName)

}

// 新增需求：箱号匹配并填写仓库代码
func processBoxNumberMatching(line int, arrivalPort string, mf_rows [][]string, awblist_rows [][]string, mf_file *excelize.File, awblist_file *excelize.File) []string {
	var Sorting_info []string
	if arrivalPort != "ATH" {
		// 定义箱号匹配列表
		type BoxInfo struct {
			airtable_code     string
			sorting_warehouse string
		}
		boxNumberMap := map[string]BoxInfo{
			"ATFA": {"AT Post Kalsdorf TEMU", "Karlsdorf"},
			"ATFY": {"AT Post Kalsdorf TEMU", "Karlsdorf"},
			"ATEA": {"AT Post Kalsdorf TEMU", "Karlsdorf"},
			"LTIC": {"AT Post Kalsdorf TEMU", "Karlsdorf"},
			"ATCD": {"AT Post Kalsdorf TEMU", "Karlsdorf"},
			"ATFT": {"AT Post Kalsdorf TEMU", "Karlsdorf"},
			"ATFH": {"AT Post Kalsdorf TEMU", "Karlsdorf"},
			"WGAT": {"AT Post Kalsdorf TEMU", "Karlsdorf"},
			"WXAT": {"AT Post Kalsdorf TEMU", "Karlsdorf"},
			"ATFL": {"AT Post Vienna TEMU", "Wien"},
			"ATFK": {"AT Post Vienna TEMU", "Wien"},
			"WRAT": {"AT Post Vienna TEMU", "Wien"},
			"WVAT": {"AT Post Vienna TEMU", "Wien"},
			"SOAT": {"AT Post Allhaming TEMU", "Allhaming"},
			"ATAY": {"AT Post Allhaming TEMU", "Allhaming"},
			"WSAT": {"AT Post Allhaming TEMU", "Allhaming"},
			"WJAT": {"AT Post Allhaming TEMU", "Allhaming"},
			"SRAT": {"AT Post Hagenbrunn TEMU", "Hagenbrunn"},
			"ATAT": {"AT Post Hagenbrunn TEMU", "Hagenbrunn"},
			"WYAT": {"AT Post Hagenbrunn TEMU", "Hagenbrunn"},
			"WNAT": {"AT Post Hagenbrunn TEMU", "Hagenbrunn"},
		}
		// 用于统计每个仓库的箱子数量
		warehouseCount := map[string]int{
			"Allhaming":  0,
			"Hagenbrunn": 0,
			"Karlsdorf":  0,
			"Wien":       0,
		}
		// 用于记录已经处理过的箱号
		processedBoxNumbers := make(map[string]bool)

		if awblist_rows[line][9] == "LGGATP-MIX" || awblist_rows[line][9] == "SOBATP-MIX" { // J 列的值是 LGGATP-MIX/SOBATP-MIX
			for i, mf_row := range mf_rows {
				if i == 0 {
					continue
				}
				boxNumber := mf_row[31] // AF 列箱号
				boxNumber_prefix := boxNumber[:4]
				if info, ok := boxNumberMap[boxNumber_prefix]; ok {
					mf_file.SetCellValue("Sheet1", "AH"+strconv.Itoa(i+1), info.airtable_code)
					if _, exists := processedBoxNumbers[boxNumber]; !exists {
						warehouseCount[info.sorting_warehouse]++
						processedBoxNumbers[boxNumber] = true
					}
				} else {
					mf_file.SetCellValue("Sheet1", "AH"+strconv.Itoa(i+1), "无匹配")
				}
			}
			/// 构建统计信息字符串
			for warehouse, count := range warehouseCount {
				Sorting_info = append(Sorting_info, fmt.Sprintf("%s: %d", strconv.Quote(warehouse), count))
			}
			//result := fmt.Sprintf("Sorting_info:{%s}", strings.Join(Sorting_info, ", "))
			// 打印统计结果
			awblist_file.SetCellValue("提单列表", "U"+strconv.Itoa(line+1), Sorting_info)

		}
		return Sorting_info
	} else {
		if awblist_rows[line][9] == "ATHGRBOX" { // J 列的值是 ATHGRBOX
			trackingNumber_prefix := mf_rows[1][28][:2] // AF 列箱号
			if trackingNumber_prefix != "76" {
				Sorting_info = []string{"BOX NOW HUB:Street Tatoiou 96, Acharne, Attica 136 72"}
				awblist_file.SetCellValue("提单列表", "U"+strconv.Itoa(line+1), Sorting_info)
			} else {
				Sorting_info = []string{"ACS  hub:Athens 36-38 Petrou Ralli Post code 12241 Athens"}
				awblist_file.SetCellValue("提单列表", "U"+strconv.Itoa(line+1), Sorting_info)
			}

		}
		return Sorting_info
	}
}

// 需求 8 和 9：发送邮件并下载提单文件
func sendEmailAndDownloadFile(update string, layout string, awb string, line int, awblist_rows [][]string, awblist_file *excelize.File) {

	flightNumber := awblist_rows[line][1]
	departurePort := awblist_rows[line][3]
	arrivalPort := awblist_rows[line][4]
	departureTime, _ := excelSerialToTime(awblist_rows[line][5])
	arrivalTime, _ := excelSerialToTime(awblist_rows[line][6])
	boxCount := awblist_rows[line][7]
	ticketCount := awblist_rows[line][10]
	billableWeight := awblist_rows[line][11]
	airtableCode := awblist_rows[line][18]
	sorting_info := awblist_rows[line][20]

	// 下载提单文件
	downloadURL := awblist_rows[line][12]
	time.Sleep(50 * time.Millisecond)
	fileName := awb + ".pdf"
	resp, _ := http.Get(downloadURL)
	defer resp.Body.Close()
	file, _ := os.Create(fileName)
	defer file.Close()
	_, _ = io.Copy(file, resp.Body)

	//创建对象
	m := mail.NewMessage()
	var from, subject, body string

	if arrivalPort == "LGG" {
		//设置发件人
		from = "LGG_Prealert<ly@ydhexpress.com>"
		//设置主题
		subject = fmt.Sprintf("%sCC AND OP INSTRUCTIONS_B2C_YDH_%s ETA %s【CT】", update, awb, arrivalTime.Format(layout))
		//设置正文
		body = fmt.Sprintf(`Dear,
Please find here enclosed the document, the customs clearance and operations instructions for below container/awb:
1. Flight Detail：%s ( %s - %s )
2. ETD ：%s (Local time)
3. ETA ：%s (Local time)
4. MAWB： %s
5. Total Bags   %s  PCS (   %s PARCELS)
6. Total Weight :   %s KG
Remark :
Delivery Channel : %s
%s`, flightNumber, departurePort, arrivalPort, departureTime.Format(layout), arrivalTime.Format(layout), awb, boxCount, ticketCount, billableWeight, airtableCode, sorting_info)
		//设置附件
		m.Attach(awb + ".pdf")
		m.Attach(awb + ".xlsx")
		m.Attach(awb + "_" + "HSH预报表格.xlsx")
		m.SetHeader("To", "ly@ydhexpress.com", "customsteam@hshworld.eu", "yangxd@ydhexpress.com", "lggops@ydhexpress.com", "szcs@ydhexpress.com")

	} else if arrivalPort == "ATH" {
		//设置发件人
		from = "ATH_Prealert<ly@ydhexpress.com>"
		//设置主题
		subject = fmt.Sprintf("%sYDH-ATH new shipment_ AWB :%s ETA %s", update, awb, arrivalTime.Format(layout))
		//设置正文
		body = fmt.Sprintf(`Dear,
Please find here enclosed the document, the customs clearance and operations instructions for below container/awb:
1. Flight Detail：%s ( %s - %s )
2. ETD ：%s (Local time)
3. ETA ：%s (Local time)
4. MAWB： %s
5. Total Bags   %s  PCS (   %s PARCELS)
6. Total Weight :   %s KG
Remark :
Delivery Channel : %s`, flightNumber, departurePort, arrivalPort, departureTime.Format(layout), arrivalTime.Format(layout), awb, boxCount, ticketCount, billableWeight, sorting_info)
		//设置附件
		m.Attach(fileName)
		m.Attach(awb + ".xlsx")
		m.SetHeader("To", "ly@ydhexpress.com", "docs@mitropouloscb.gr", "yangxd@ydhexpress.com", "lggops@ydhexpress.com", "szcs@ydhexpress.com", "daisy.european@gmail.com")

	} else if arrivalPort == "SOB" {
		//设置发件人
		from = "SOB_Prealert<ly@ydhexpress.com>"
		//设置主题
		subject = fmt.Sprintf("%sYDH-SOB new shipment_ AWB :%s ETA %s", update, awb, arrivalTime.Format(layout))
		//设置正文
		body = fmt.Sprintf(`Dear,
Please find here enclosed the document, the customs clearance and operations instructions for below container/awb:
1. Flight Detail：%s ( %s - %s )
2. ETD ：%s (Local time)
3. ETA ：%s (Local time)
4. MAWB： %s
5. Total Bags   %s  PCS (   %s PARCELS)
6. Total Weight :   %s KG
Remark :
Delivery Channel : %s
%s`, flightNumber, departurePort, arrivalPort, departureTime.Format(layout), arrivalTime.Format(layout), awb, boxCount, ticketCount, billableWeight, airtableCode, sorting_info)
		//设置附件
		m.Attach(fileName)
		m.Attach(awb + ".xlsx")
		m.SetHeader("To", "ly@ydhexpress.com", "customs@i-cont.eu", "yangxd@ydhexpress.com", "lggops@ydhexpress.com", "szcs@ydhexpress.com", "aircargo@ttdparcel.eu", "ttd-bot@outlook.com", "zsherwin@ydhexpress.com")

	} else if arrivalPort == "BEG" {
		//设置发件人
		from = "BEG_Prealert<ly@ydhexpress.com>"
		//设置主题
		subject = fmt.Sprintf("%sYDH-PRE ALERT - AWB#%s ETA %s", update, awb, arrivalTime.Format(layout))
		//设置正文
		body = fmt.Sprintf(`Dear,
Please find here enclosed the document, the customs clearance and operations instructions for below container/awb:
1. Flight Detail：%s ( %s - %s )
2. ETD ：%s (Local time)
3. ETA ：%s (Local time)
4. MAWB： %s
5. Total Bags   %s  PCS (   %s PARCELS)
6. Total Weight :   %s KG
`, flightNumber, departurePort, arrivalPort, departureTime.Format(layout), arrivalTime.Format(layout), awb, boxCount, ticketCount, billableWeight)
		//设置附件
		m.Attach(awb + ".pdf")
		m.Attach(awb + ".xlsx")
		m.SetHeader("Cc", "ly@ydhexpress.com",
			"yangxd@ydhexpress.com",
			"lggops@ydhexpress.com",
			"cmyd@ydhexpress.com ",
			"cp@ydhexpress.com",
			"cs@ydhexpress.com",
			"zqq@ydhexpress.com",
			"duanhaoying@ydhexpress.com",
			"cyj@ydhexpress.com",
			"yingfucaiwu@ydhexpress.com",
			"xiaoyudian@ydhexpress.com",
			"szcs@ydhexpress.com")
		m.SetHeader("To", "nikola.petrovic@colisped.rs", "vladimir.kostadinovic@colisped.rs", "jana.kostadinovic@colisped.rs")
	} else if arrivalPort == "OTP" {
		//设置发件人
		from = "OTP_Prealert<ly@ydhexpress.com>"
		//设置主题
		subject = fmt.Sprintf("%sYDH-OTP new shipment_ AWB :%s ETA %s", update, awb, arrivalTime.Format(layout))
		//设置正文
		body = fmt.Sprintf(`Dear,
Please find here enclosed the document, the customs clearance and operations instructions for below container/awb:
1. Flight Detail：%s ( %s - %s )
2. ETD ：%s (Local time)
3. ETA ：%s (Local time)
4. MAWB： %s
5. Total Bags   %s  PCS (   %s PARCELS)
6. Total Weight :   %s KG
`, flightNumber, departurePort, arrivalPort, departureTime.Format(layout), arrivalTime.Format(layout), awb, boxCount, ticketCount, billableWeight)
		//设置附件
		m.Attach(awb + ".pdf")
		m.Attach(awb + ".xlsx")
		m.SetHeader("Cc", "wanglei@ydhexpress.com",
			"yangxd@ydhexpress.com",
			"lggops@ydhexpress.com",
			"szcs@ydhexpress.com",
			"ly@ydhexpress.com")
		m.SetHeader("To", "dan.belcea@asgwind.eu", "adrian@asgwind.eu", "cristinac@asgwind.eu")
	}
	//fmt.Println(subject)
	//fmt.Println(body)
	// 配置 SMTP 服务器信息
	username := "ly@ydhexpress.com"
	password := "fjsoTBLpm4iahBPK"

	m.SetHeader("From", from)
	//m.SetHeader("Cc", "ly@ydhexpress.com")
	//m.SetHeader("To", "ly@ydhexpress.com")

	m.SetHeader("Subject", subject)
	m.SetBody("text/plain", body)

	//建立链接&发送
	d := mail.NewDialer("smtp.exmail.qq.com", 465, username, password)
	time.Sleep(50 * time.Millisecond)
	if err := d.DialAndSend(m); err != nil {
		log.Printf("发送失败: %v", err)
	}
	log.Print(awb, "邮件已发送√")

}

// 匹配 HS Code
func matchHSCode(hs_column int, description_column int, mf_file *excelize.File, mf_rows [][]string) {
	hs_column_name := numberToExcelColumn(hs_column)
	description_column_name := numberToExcelColumn(description_column)

	replacements := map[string]string{
		"Medical ":             "",
		"medical ":             "",
		"cigarette":            "",
		"Cigarette":            "",
		"Excipients":           "pipe",
		"gun":                  "",
		"hand shot":            "simulation toy",
		"Night scope":          "Telescope",
		"cellophane":           "simulation toy",
		"antipyretic treasure": "pipe",
		"insect":               "",
		"Breathalyzer":         "instrument parts",
		"organic surfactant":   "make up",
		"cat ":                 "",
		"respirator":           "machine",
		"icide":                "decoration parts",
		"Stuffed Animal":       "Stuffed toy",
		"Fireworks":            "decoration parts",
		"Building block":       "brick part",
		"Flood":                "",
		"Construction":         "Component",
		"helmet":               "plastic hat",
		"Soil tester":          "tester",
		"Perfume":              "make up",
		"perfume":              "make up",
		"Baby ma3 jia3":        "suit",
		"toy":                  "for fun",
		"Toy":                  "for fun",
		"Gift":                 "",
		"doll":                 "for play",
		"Doll":                 "for play",
		"Building block toy":   "piece Component",
	}
	var wupipei int8
	// 打开 ku 文件
	file_hscode, _ := excelize.OpenFile("ku.xlsx")
	pinmingkuRows := file_hscode.GetRows("品名库")

	for i, mf_row := range mf_rows {
		// 跳过表头
		if i == 0 {
			continue
		}
		found := false

		mf_hscode := mf_row[hs_column-1]          // Y 列 HS CODE
		mf_enname := mf_row[description_column-1] // Z 列品名

		for oldWord, newWord := range replacements {
			mf_enname = strings.Replace(mf_enname, oldWord, newWord, -1)
			mf_file.SetCellValue("Sheet1", description_column_name+strconv.Itoa(i+1), mf_enname)
		}

		// 判断 8/73 开头
		if string(mf_hscode)[0:1] == "8" {
			found = true
			continue
		} else if string(mf_hscode)[:2] == "73" {
			mf_file.SetCellValue("Sheet1", hs_column_name+strconv.Itoa(i+1), "9607190000")
			found = true
			continue
		}
		for j, pinmingkuRow := range pinmingkuRows {
			// 跳过品名库的表头
			if j == 0 {
				continue
			}
			// 获取品名库的英文品名,hscode10 位，hdcode6 位
			pinmingkuName := pinmingkuRow[0]
			pinmingku_hs_10 := pinmingkuRow[1]
			pinmingku_hs_6 := pinmingkuRow[2]

			if strings.ToUpper(mf_enname) == strings.ToUpper(pinmingkuName) {
				mf_file.SetCellValue("Sheet1", hs_column_name+strconv.Itoa(i+1), pinmingku_hs_10)
				found = true
				break
			} else if mf_hscode[:6] == (pinmingku_hs_6) {
				found = true
				break
			} else if strings.Contains(strings.ToUpper(pinmingkuName), strings.ToUpper(mf_enname)) || strings.Contains(strings.ToUpper(mf_enname), strings.ToUpper(pinmingkuName)) {
				mf_file.SetCellValue("Sheet1", hs_column_name+strconv.Itoa(i+1), pinmingku_hs_10)
				found = true
				break
			}
		}

		if !found {
			mf_row[24] = "无匹配"
			mf_file.SetCellValue("Sheet1", hs_column_name+strconv.Itoa(i+1), "无匹配")
			wupipei++
		}
	}
	// 如果无匹配的数量超过 5 个，则将所有无匹配的值都改为 "3926200000"
	if wupipei <= 5 {
		for i, row := range mf_rows {
			if i == 0 {
				continue
			}
			if row[24] == "无匹配" {
				mf_file.SetCellValue("Sheet1", hs_column_name+strconv.Itoa(i+1), "3926200000")
			}
		}
	} else {
		fmt.Println("无匹配个数已超过 5 个，请运行完毕后手工匹配")
		mf_file.Save()
		time.Sleep(99999 * time.Hour)

	}
	file_hscode.Save()
}

func main() {
	var update = ""

	var AWBListFilePath = []string{
		"C:\\Users\\Administrator\\Nutstore\\1\\数据源\\TEMU\\temu LGG清关\\LGG提单列表.xlsx",
		"C:\\Users\\Administrator\\Nutstore\\1\\数据源\\TEMU\\temu 非LGG清关\\除LGG外的提单列表.xlsx",
		"F:\\数据源\\TEMU\\temu LGG清关\\LGG提单列表.xlsx",
		"F:\\数据源\\TEMU\\temu 非LGG清关\\除LGG外的提单列表.xlsx",
	}
	fmt.Println("当前提单列表文件路径：", AWBListFilePath)
	// 提示用户是否要进行操作
	fmt.Println("1. 修改提单列表文件路径")
	fmt.Println("2. 更新Update邮件")
	fmt.Print("请选择要进行的操作，若不输入，默认不进行任何操作：")

	inputScanner := bufio.NewScanner(os.Stdin)
	inputScanner.Scan()
	user_input := strings.TrimSpace(strings.ToUpper(inputScanner.Text()))

	//解析用户选择
	switch user_input {
	case "1":
		// 提示用户输入第一个文件路径
		fmt.Print("请输入LGG提单列表文件路径: ")
		scanner := bufio.NewScanner(os.Stdin)
		if scanner.Scan() {
			AWBListFilePath[0] = scanner.Text()
		}

		// 提示用户输入第二个文件路径
		fmt.Print("请输入非LGG提单列表文件路径: ")
		if scanner.Scan() {
			AWBListFilePath[1] = scanner.Text()
		}

		// 打印用户输入的文件路径
		fmt.Println("你输入的文件路径为:", AWBListFilePath)
	case "2":
		update = "(update)"
	default:
		//fmt.Println("未选择，不进行任何操作。")
	}

	var layout string
	layout = "2006-01-02 15:04:05"

	// 输入提单号
	//var awb string
	//fmt.Print("请输入提单号：")
	//fmt.Scanln(&awb)
	// 输入多个提单号，支持空格、逗号、换行分隔
	fmt.Print("请输入提单号，多个提单号可以用空格、逗号或换行分隔，输入完成后按 Ctrl + D结束输入：")
	scanner := bufio.NewScanner(os.Stdin)
	var input string
	for scanner.Scan() {
		line := scanner.Text()
		input += line + " "
	}

	// 使用正则表达式分割输入，匹配逗号、空格或换行符
	re := regexp.MustCompile(`[,\s]+`)
	awbs := re.Split(strings.TrimSpace(input), -1)
	log.Printf("待发送的提单号列表: %v", awbs)

	for _, awb := range awbs {
		awb = strings.TrimSpace(awb)
		if awb == "" {
			continue
		}
		log.Print("正在处理的提单号是：", awb)
		var line int
		var arrivalPort string

		// 打开 manifest 文件并读取各列数据
		mf_file, _ := excelize.OpenFile(awb + ".xlsx") //manifest文件
		mf_rows := mf_file.GetRows("Sheet1")           //manifest行

		// 遍历文件路径列表，依次搜索提单号
		var awblist_file *excelize.File
		var awblist_rows [][]string
		for _, filePath := range AWBListFilePath {
			// 检查文件是否存在
			if _, err := os.Stat(filePath); os.IsNotExist(err) {
				//log.Printf("文件 %s 不存在，跳过该文件", filePath)
				continue
			}
			awblist_file, _ = excelize.OpenFile(filePath) // 提单列表文件
			awblist_rows = awblist_file.GetRows("提单列表")   // 提单列表行
			for k, awblist_row := range awblist_rows {
				if k == 0 {
					continue
				}
				if awblist_row[0] == awb {
					line = k                     //对应行
					arrivalPort = awblist_row[4] //港口
					break
				}
			}
			time.Sleep(5 * time.Millisecond)
			if line > 0 {
				break

			}

		}
		if line < 1 {
			log.Println(awb, "提单号不存在")
			time.Sleep(99999 * time.Hour)

		}

		if arrivalPort == "LGG" || arrivalPort == "ATH" || arrivalPort == "SOB" {
			// 处理需求  删除最后一行
			processLastRow(20, mf_rows, mf_file)
			mf_rows = mf_file.GetRows("Sheet1")
			// 匹配 HS Code
			matchHSCode(25, 26, mf_file, mf_rows)
			//处理需求 1： 删除客户单号列
			deleteADColumn(30, mf_file)
			// 处理需求 2 过滤AC列的%,填补consignor city & post code : zhaoqing, 526200
			processACColumn(arrivalPort, mf_file, mf_rows)
			// 处理需求 3  处理 M 列和 N 列 填补空白买家信息，截取前50字符，地址没有门牌号就补0
			processMNColumns(13, 14, 29, mf_file, mf_rows)
			mf_rows = mf_file.GetRows("Sheet1")
			//奥地利分仓信息
			processBoxNumberMatching(line, arrivalPort, mf_rows, awblist_rows, mf_file, awblist_file)
			awblist_file.Save()
			mf_rows = mf_file.GetRows("Sheet1")
			// 处理需求 4 处理150欧
			processAmountLimit(13, 14, 21, 29, mf_file, mf_rows)
			// 处理需求 5 处理提单列表的重量分摊
			processLGGFile(23, line, mf_rows, awblist_rows, mf_file)
			// 需求需求 7：处理 HSH 预报表格.xlsx
			processHSHFile(layout, awb, line, awblist_rows)
			mf_file.Save()
			// 需求 8 和 9：发送邮件并下载提单文件
			awblist_rows = awblist_file.GetRows("提单列表")
			sendEmailAndDownloadFile(update, layout, awb, line, awblist_rows, awblist_file)
			// 保存工作簿文件
			mf_file.Save()
		} else if arrivalPort == "BEG" {
			rate := FetchExchangeRate()
			fmt.Println(rate)
			// 处理需求  删除最后一行
			processLastRow(2, mf_rows, mf_file)
			mf_rows = mf_file.GetRows("Sheet1")
			// 处理需求 4 处理50欧
			BEGprocessAmountLimit(mf_file, mf_rows)
			// 保存工作簿文件
			mf_file.Save()
			awblist_rows = awblist_file.GetRows("提单列表")
			sendEmailAndDownloadFile(update, layout, awb, line, awblist_rows, awblist_file)
			// 保存工作簿文件
			mf_file.Save()
		} else if arrivalPort == "OTP" {
			// 处理需求  删除最后一行
			processLastRow(20, mf_rows, mf_file)
			mf_rows = mf_file.GetRows("Sheet1")
			// 处理需求 3  处理 M 列和 N 列 填补空白买家信息，截取前50字符，地址没有门牌号就补0
			processMNColumns(14, 15, 1, mf_file, mf_rows)
			// 匹配 HS Code
			matchHSCode(19, 6, mf_file, mf_rows)
			// 处理需求 4 处理150欧
			processAmountLimit(14, 15, 23, 1, mf_file, mf_rows)
			//填补zhaoqing & zip code
			processOTPACColumn(arrivalPort, mf_file, mf_rows)
			// 处理需求 5 处理提单列表的重量分摊
			processOTPFile(5, line, mf_rows, awblist_rows, mf_file)
			// 保存工作簿文件
			mf_file.Save()
			// 需求 8 和 9：发送邮件并下载提单文件
			awblist_rows = awblist_file.GetRows("提单列表")
			sendEmailAndDownloadFile(update, layout, awb, line, awblist_rows, awblist_file)
			// 保存工作簿文件
			mf_file.Save()
		}

	}
	fmt.Println("全部发送完成,可关闭窗口")
	time.Sleep(99999 * time.Hour)

}

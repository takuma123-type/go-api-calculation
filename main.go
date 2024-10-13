package main

import (
	"fmt"
	"log"
	"math"
	"strconv"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
	"gonum.org/v1/gonum/graph"
	"gonum.org/v1/gonum/graph/simple"
)

// 駅間のエッジを追加
func addEdgesWithDistances(G *simple.WeightedUndirectedGraph, stations []string, distances map[string]float64) {
	nodes := make(map[string]graph.Node)

	// 駅ごとのノードを追加
	for _, station := range stations {
		node := G.NewNode()
		G.AddNode(node)
		nodes[station] = node
	}

	// 駅間エッジを追加
	for i, station1 := range stations {
		station2Inner := stations[(i+1)%len(stations)]               // 内回り
		station2Outer := stations[(i-1+len(stations))%len(stations)] // 外回り

		// 内回りと外回りのエッジを追加
		if station1 != station2Inner {
			if dist, ok := distances[station1+"-"+station2Inner]; ok {
				G.SetWeightedEdge(G.NewWeightedEdge(nodes[station1], nodes[station2Inner], dist))
			}
		}

		if station1 != station2Outer {
			if dist, ok := distances[station2Outer+"-"+station1]; ok {
				G.SetWeightedEdge(G.NewWeightedEdge(nodes[station1], nodes[station2Outer], dist))
			}
		}
	}
}

// 内回りと外回りのルートを評価して最短距離を選択
func getShortestRoute(station1, station2 string, distances map[string]float64) float64 {
	innerKey := station1 + "-" + station2
	outerKey := station2 + "-" + station1

	innerDistance, innerExists := distances[innerKey]
	outerDistance, outerExists := distances[outerKey]

	if innerExists && outerExists {
		return math.Min(innerDistance, outerDistance)
	} else if innerExists {
		return innerDistance
	} else if outerExists {
		return outerDistance
	}
	return -1 // 駅間の距離が存在しない場合、-1を返す（エラー検知用）
}

// エレベータの位置情報を活用してホーム内移動距離を計算
func calculateHomeDistances(startStation, endStation string, startDir, endDir string, elevatorInfo map[string]map[string]string) float64 {
	startElevators := elevatorInfo[startStation]
	endElevators := elevatorInfo[endStation]

	// 最短距離を計算
	minDistance := math.MaxFloat64
	for startCar, startStatus := range startElevators {
		if startStatus == "Y" || startStatus == "e" {
			startCarNum, _ := strconv.Atoi(startCar)
			for endCar, endStatus := range endElevators {
				if endStatus == "Y" || endStatus == "e" {
					endCarNum, _ := strconv.Atoi(endCar)
					// 同じ方向のエレベータだけを考慮
					if startDir == endDir {
						distance := math.Abs(float64(startCarNum - endCarNum))
						if distance < minDistance {
							minDistance = distance
						}
					}
				}
			}
		}
	}
	// 最短距離が更新されていない場合、エラーチェック
	if minDistance == math.MaxFloat64 {
		return -1 // エラーとして -1 を返す
	}
	return minDistance
}

// Excelデータを読み込む
func readExcelData(inputFile string) ([]string, map[string]float64, map[string]map[string]string, int, []string, error) {
	f, err := excelize.OpenFile(inputFile)
	if err != nil {
		return nil, nil, nil, 0, nil, err
	}

	// 駅リストの取得
	stations := []string{}
	rows, err := f.GetRows("内回り")
	if err != nil {
		return nil, nil, nil, 0, nil, err
	}
	for _, row := range rows {
		if len(row) > 0 {
			stations = append(stations, row[0])
		}
	}

	// 駅間距離の取得
	distances := make(map[string]float64)
	rows, err = f.GetRows("距離データ")
	if err != nil {
		return nil, nil, nil, 0, nil, err
	}
	for _, row := range rows {
		if len(row) >= 3 {
			station1 := row[0]
			station2 := row[1]
			distance, _ := strconv.ParseFloat(row[2], 64)
			distances[station1+"-"+station2] = distance
		}
	}

	// エレベータの情報（増設可能 or 既存エレベータ）と最大増設数の計算
	elevatorInfo := make(map[string]map[string]string)
	maxInstallations := 0
	var possibleInstallations []string
	for _, sheetName := range []string{"内回り", "外回り"} {
		rows, err := f.GetRows(sheetName)
		if err != nil {
			return nil, nil, nil, 0, nil, err
		}
		for _, row := range rows {
			station := row[0]
			if elevatorInfo[station] == nil {
				elevatorInfo[station] = make(map[string]string)
			}
			for i, status := range row[1:] {
				carNumber := strconv.Itoa(i + 1)
				elevatorInfo[station][carNumber] = status
				if status == "Y" {
					possibleInstallations = append(possibleInstallations, fmt.Sprintf("%s_%s_%s", station, carNumber, sheetName))
					maxInstallations++
				}
			}
		}
	}

	return stations, distances, elevatorInfo, maxInstallations, possibleInstallations, nil
}

// 結果を保存（増設数ごとのシートも作成）
func saveResults(finalOutputFile string, optimalResults []map[string]interface{}, detailedResults map[int][]map[string]interface{}) {
	f := excelize.NewFile()

	// 最適結果のシート作成
	sheetIndex := f.NewSheet("最適結果")
	f.SetActiveSheet(sheetIndex)

	// ヘッダーを設定
	headers := []string{"増設数", "増設箇所", "駅間移動距離", "ホーム内移動距離", "評価ペア数", "1ペアあたりホーム内移動距離", "改善率"}
	for i, header := range headers {
		cell, _ := excelize.CoordinatesToCellName(i+1, 1)
		f.SetCellValue("最適結果", cell, header)
	}

	// データを入力
	for i, result := range optimalResults {
		for j, key := range headers {
			cell, _ := excelize.CoordinatesToCellName(j+1, i+2)
			f.SetCellValue("最適結果", cell, result[key])
		}
	}

	// 増設数ごとのシートに詳細結果を保存
	for count, results := range detailedResults {
		sheetName := fmt.Sprintf("増設数_%d", count)
		f.NewSheet(sheetName)

		// ヘッダーを設定
		for i, header := range headers {
			cell, _ := excelize.CoordinatesToCellName(i+1, 1)
			f.SetCellValue(sheetName, cell, header)
		}

		// データを入力
		for i, result := range results {
			for j, key := range headers {
				cell, _ := excelize.CoordinatesToCellName(j+1, i+2)
				f.SetCellValue(sheetName, cell, result[key])
			}
		}
	}

	if err := f.SaveAs(finalOutputFile); err != nil {
		log.Fatalf("結果の保存に失敗しました: %v", err)
	}
	fmt.Printf("最終結果を保存しました: %s\n", finalOutputFile)
}

func main() {
	inputFile := "yamanotemodel.xlsx"

	// Excelデータの読み込み
	stations, distances, elevatorInfo, maxInstallations, possibleInstallations, err := readExcelData(inputFile)
	if err != nil {
		log.Fatalf("Excelファイルの読み込みに失敗しました: %v", err)
	}

	// グラフを初期化
	G := simple.NewWeightedUndirectedGraph(0, 0)
	addEdgesWithDistances(G, stations, distances)

	// 詳細結果を保存するためのマップ
	detailedResults := make(map[int][]map[string]interface{})

	// 増設数に応じてエレベータ設置場所を評価
	optimalResults := []map[string]interface{}{}
	for count := 0; count <= maxInstallations; count++ {
		fmt.Printf("Evaluating %d elevators\n", count)

		// 各駅間の距離を計算
		totalDistance := 0.0
		totalHomeDistance := 0.0
		evaluatedPairs := 0

		for i := 0; i < len(stations); i++ {
			for j := i + 1; j < len(stations); j++ {
				station1 := stations[i]
				station2 := stations[j]

				// 駅間移動距離の計算
				stationDistance := getShortestRoute(station1, station2, distances)
				if stationDistance > 0 {
					totalDistance += stationDistance
				}

				// ホーム内移動距離の計算
				homeDistance := calculateHomeDistances(station1, station2, "内回り", "外回り", elevatorInfo)
				if homeDistance > 0 {
					totalHomeDistance += homeDistance
				}
				evaluatedPairs++
			}
		}

		// 1ペアあたりのホーム内移動距離
		homeDistancePerPair := totalHomeDistance / float64(evaluatedPairs)

		// 結果を保存
		result := map[string]interface{}{
			"増設数":                       count,
			"増設箇所":                     strings.Join(possibleInstallations[:count], ","),
			"駅間移動距離":                 totalDistance,
			"ホーム内移動距離":             totalHomeDistance,
			"評価ペア数":                   evaluatedPairs,
			"1ペアあたりホーム内移動距離": homeDistancePerPair,
			"改善率":                       10.0, // 仮の値
		}
		optimalResults = append(optimalResults, result)
		detailedResults[count] = append(detailedResults[count], result)
	}

	// 結果をExcelファイルに保存
	finalOutputFile := fmt.Sprintf("final_results_%s.xlsx", time.Now().Format("20060102_150405"))
	saveResults(finalOutputFile, optimalResults, detailedResults)
}

// Simulador de inventario para el problema 5.5.
// Consume exclusivamente una secuencia externa de uniformes R_i.
package main

import (
	"flag"
	"fmt"
	"math"
	"os"
	"strconv"
	"strings"
)

type Params struct {
	Days         int
	DemandN      int
	DemandP      float64
	LeadLambda   float64
	HoldingCost  float64
	ShortageCost float64
	OrderCost    float64
	InitialStock int
}
type Policy struct {
	Name                    string
	OrderUpTo, ReorderPoint int
	ReviewDaily             bool
}
type Result struct {
	Total, Holding, Shortage, Ordering float64
	Orders, Demand, ShortageUnits      int
}
type uniformCursor struct {
	values []float64
	index  int
}

func (c *uniformCursor) next() (float64, error) {
	if c.index >= len(c.values) {
		return 0, fmt.Errorf("la secuencia se agotó después de %d valores", len(c.values))
	}
	value := c.values[c.index]
	c.index++
	return value, nil
}

func sampleBinomial(c *uniformCursor, n int, p float64) (int, error) {
	value := 0
	for range n {
		u, err := c.next()
		if err != nil {
			return 0, err
		}
		if u < p {
			value++
		}
	}
	return value, nil
}

// Inverse CDF: cada plazo utiliza exactamente un R_i entregado por el usuario.
func samplePoisson(c *uniformCursor, lambda float64) (int, error) {
	u, err := c.next()
	if err != nil {
		return 0, err
	}
	probability, cumulative, value := math.Exp(-lambda), math.Exp(-lambda), 0
	for u > cumulative {
		value++
		probability *= lambda / float64(value)
		cumulative += probability
	}
	return value, nil
}

func simulate(p Params, policy Policy, uniforms []float64) (Result, error) {
	cursor := uniformCursor{values: uniforms}
	inventory, deliveries := p.InitialStock, make(map[int]int)
	var result Result
	for day := 1; day <= p.Days; day++ {
		inventory += deliveries[day]
		delete(deliveries, day)
		demand, err := sampleBinomial(&cursor, p.DemandN, p.DemandP)
		if err != nil {
			return Result{}, fmt.Errorf("día %d, demanda: %w", day, err)
		}
		result.Demand += demand
		inventory -= demand
		if inventory >= 0 {
			result.Holding += float64(inventory) * p.HoldingCost
		} else {
			result.Shortage += float64(-inventory) * p.ShortageCost
			result.ShortageUnits += -inventory
		}
		if (policy.ReviewDaily || inventory <= policy.ReorderPoint) && inventory < policy.OrderUpTo {
			leadTime, err := samplePoisson(&cursor, p.LeadLambda)
			if err != nil {
				return Result{}, fmt.Errorf("día %d, plazo de entrega: %w", day, err)
			}
			deliveries[day+max(1, leadTime)] += policy.OrderUpTo - inventory
			result.Ordering += p.OrderCost
			result.Orders++
		}
	}
	result.Total = result.Holding + result.Shortage + result.Ordering
	return result, nil
}

func loadUniforms(path string) ([]float64, error) {
	content, err := os.ReadFile(path)
	if err != nil {
		return nil, err
	}
	tokens := strings.FieldsFunc(string(content), func(r rune) bool { return r == ',' || r == ';' || r == '\n' || r == '\r' || r == '\t' || r == ' ' })
	values := make([]float64, 0, len(tokens))
	for _, token := range tokens {
		if strings.EqualFold(token, "ri") || strings.EqualFold(token, "r_i") {
			continue
		}
		value, err := strconv.ParseFloat(strings.ReplaceAll(token, ",", "."), 64)
		if err != nil || value < 0 || value >= 1 {
			return nil, fmt.Errorf("%q no es un R_i válido; se espera 0 <= R_i < 1", token)
		}
		values = append(values, value)
	}
	if len(values) == 0 {
		return nil, fmt.Errorf("el archivo no contiene R_i válidos")
	}
	return values, nil
}

func printResult(policy Policy, r Result, days int) {
	fmt.Printf("%s\n", policy.Name)
	fmt.Printf("  costo total:            $%10.2f  ($%.2f/día)\n", r.Total, r.Total/float64(days))
	fmt.Printf("  mantenimiento:          $%10.2f\n", r.Holding)
	fmt.Printf("  faltantes:              $%10.2f\n", r.Shortage)
	fmt.Printf("  órdenes:                $%10.2f  (%d órdenes)\n", r.Ordering, r.Orders)
}

func main() {
	p := Params{}
	var uniformFile string
	flag.IntVar(&p.Days, "dias", 0, "días de la corrida")
	flag.StringVar(&uniformFile, "uniformes", "", "archivo CSV/TXT con R_i externos")
	flag.IntVar(&p.InitialStock, "inventario-inicial", 8, "inventario neto inicial")
	flag.Parse()
	if p.Days <= 0 || uniformFile == "" {
		fmt.Println("Uso: simulador -dias 285 -uniformes ruta/a/ri.csv")
		return
	}
	uniforms, err := loadUniforms(uniformFile)
	if err != nil {
		fmt.Printf("No se pudieron leer los R_i: %v\n", err)
		return
	}
	if p.Days*7 > len(uniforms) {
		fmt.Printf("Se requieren hasta %d R_i para %d días; el archivo contiene %d.\n", p.Days*7, p.Days, len(uniforms))
		return
	}
	p.DemandN, p.DemandP, p.LeadLambda = 6, 0.5, 3
	p.HoldingCost, p.ShortageCost, p.OrderCost = 1, 10, 50
	daily := Policy{"Política 1: revisar cada día, subir a 8", 8, 0, true}
	reorder := Policy{"Política 2: cuando inventario <= 10, subir a 30", 30, 10, false}
	r1, err := simulate(p, daily, uniforms)
	if err != nil {
		fmt.Printf("Error en política 1: %v\n", err)
		return
	}
	r2, err := simulate(p, reorder, uniforms)
	if err != nil {
		fmt.Printf("Error en política 2: %v\n", err)
		return
	}
	fmt.Printf("Corrida de %d días usando %d R_i externos de %s\n\n", p.Days, len(uniforms), uniformFile)
	printResult(daily, r1, p.Days)
	printResult(reorder, r2, p.Days)
	if r1.Total < r2.Total {
		fmt.Printf("\nMás económica: Política 1 (ahorro $%.2f).\n", r2.Total-r1.Total)
	} else {
		fmt.Printf("\nMás económica: Política 2 (ahorro $%.2f).\n", r1.Total-r2.Total)
	}
}

package main

import (
	"errors"
	"strconv"
)

type Yml_catalog struct {
	Shop Shop `xml:"shop"`
}

type Shop struct {
	Name                  string        `xml:"name"`
	Company               string        `xml:"company"`
	Url                   string        `xml:"url"`
	Enable_auto_discounts string        `xml:"enable_auto_discounts"`
	Currencies            Currencies    `xml:"currencies"`
	Categories            Categories    `xml:"categories"`
	Delivery_opts         Delivery_opts `xml:"delivery-options"`
	Offers                Offers        `xml:"offers"`
}

type Currencies struct {
	Currencies []Currency `xml:"currency"`
}

type Currency struct {
	Id   string `xml:"id,attr"`
	Rate string `xml:"rate,attr"`
}

type Categories struct {
	Categories []Category `xml:"category"`
}

type Category struct {
	Id       string `xml:"id,attr"`
	ParentId string `xml:"parentId,attr"`
	Name     string `xml:",chardata"`
}

type Delivery_opts struct {
	Options []Option `xml:"option"`
}

type Option struct {
	Cost         string `xml:"cost,attr"`
	Days         string `xml:"days,attr"`
	Order_before string `xml:"order-before,attr"`
}

type Offers struct {
	Offers []Offer `xml:"offer"`
}

type Offer struct {
	Id            string        `xml:"id,attr"`
	Available     string        `xml:"available,attr"`
	Url           string        `xml:"url"`
	Price         string        `xml:"price"`
	CurrencyId    string        `xml:"currencyId"`
	CategoryId    string        `xml:"categoryId"`
	Picture       string        `xml:"picture"`
	Store         string        `xml:"store"`
	Pickup        string        `xml:"pickup"`
	Delivery      string        `xml:"delivery"`
	Delivery_opts Delivery_opts `xml:"delivery-options"`
	Name          string        `xml:"name"`
	Vendor        string        `xml:"vendor"`
	VendorCode    string        `xml:"vendorCode"`
	Weight        string        `xml:"weight"`
	Description   string        `xml:"description"`
	Sales_notes   string        `xml:"sales_notes"`
	Barcodes      []string      `xml:"barcode"`
	TypePrefix    string        `xml:"typePrefix"`
	Dimensions    string        `xml:"dimensions"`
	Model         string        `xml:"model"'`
	Params        []Param       `xml:"param"`
}

type Param struct {
	Name  string `xml:"name,attr"`
	Unit  string `xml:"unit,attr"`
	Value string `xml:",chardata"`
}

//get table with offer's linear field
func (c *Yml_catalog) GetOfferTable() Table {
	table := Table{}
	categoryNameMap := make(map[string]string)
	for _, category := range c.Shop.Categories.Categories {
		categoryNameMap[category.Id] = category.Name
	}
	bcCount := 0
	for _, offer := range c.Shop.Offers.Offers {
		if bcCount < len(offer.Barcodes) {
			bcCount = len(offer.Barcodes)
		}
	}

	for _, offer := range c.Shop.Offers.Offers {
		table.AddRow()
		table.SetCellValue("Id", offer.Id)
		table.SetCellValue("Available", offer.Available)
		table.SetCellValue("Url", offer.Url)
		table.SetCellValue("Price", offer.Price)
		table.SetCellValue("CurrencyID", offer.CurrencyId)
		table.SetCellValue("CategoryId", offer.CategoryId)
		table.SetCellValue("CategoryName", categoryNameMap[offer.CategoryId])
		table.SetCellValue("Picture", offer.Picture)
		table.SetCellValue("Store", offer.Store)
		table.SetCellValue("Pickup", offer.Pickup)
		table.SetCellValue("Delivery", offer.Delivery)
		table.SetCellValue("Name", offer.Name)
		table.SetCellValue("Vendor", offer.Vendor)
		table.SetCellValue("VendorCode", offer.VendorCode)
		table.SetCellValue("Weight", offer.Weight)
		table.SetCellValue("Description", offer.Description)
		table.SetCellValue("Sales_notes", offer.Sales_notes)
		for i := 0; i < bcCount; i++ {
			if i < len(offer.Barcodes) {
				table.SetCellValue("Barcode"+strconv.Itoa(i+1), offer.Barcodes[i])
			} else {
				table.SetCellValue("Barcode"+strconv.Itoa(i+1), "")
			}

		}

		table.SetCellValue("TypePrefix", offer.TypePrefix)
		table.SetCellValue("dimensions", offer.Dimensions)
		table.SetCellValue("Model", offer.Model)
	}
	return table
}

func (c *Yml_catalog) GetCategoryTable() Table {
	table := Table{}
	for _, category := range c.Shop.Categories.Categories {
		table.AddRow()
		table.SetCellValue("Id", category.Id)
		table.SetCellValue("ParendId", category.ParentId)
		table.SetCellValue("Name", category.Name)
	}
	return table
}

func (c *Yml_catalog) GetCategoryTreeTable() Table {
	table := Table{}

	parentMap := make(map[string]string)
	parentsMap := make(map[string][]string)
	hasChildMap := make(map[string]bool)
	nameMap := make(map[string]string)
	for _, category := range c.Shop.Categories.Categories {
		parentMap[category.Id] = category.ParentId
		nameMap[category.Id] = category.Name
		if category.ParentId != "" {
			hasChildMap[category.ParentId] = true
		}
	}
	maxLevel := 1
	for _, category := range c.Shop.Categories.Categories {
		tempSlice := make([]string, 0)
		tempCategoryId := category.Id
		tempSlice = append(tempSlice, tempCategoryId)
		for parentMap[tempCategoryId] != "" {
			tempSlice = append(tempSlice, parentMap[tempCategoryId])
			tempCategoryId = parentMap[tempCategoryId]
		}
		parentsMap[category.Id] = tempSlice
		if maxLevel < len(tempSlice) {
			maxLevel = len(tempSlice)
		}

	}
	for i := 0; i < maxLevel; i++ {
		table.AddColumn("ID " + strconv.Itoa(i+1))
		table.AddColumn("name " + strconv.Itoa(i+1))
	}
	for _, category := range c.Shop.Categories.Categories {
		if _, ok := hasChildMap[category.Id]; !ok {
			table.AddRow()
			tempSlice := parentsMap[category.Id]
			for i := 0; i < len(tempSlice); i++ {
				table.SetCellValue("ID "+strconv.Itoa(len(tempSlice)-i), tempSlice[i])
				table.SetCellValue("name "+strconv.Itoa(len(tempSlice)-i), nameMap[tempSlice[i]])
			}

		}
	}

	return table
}

//fill Table object by params from offers
func (c *Yml_catalog) GetParamsTable() Table {
	table := Table{}
	for _, offer := range c.Shop.Offers.Offers {
		for _, param := range offer.Params {
			err := table.AddRow()
			if err != nil {
				return table
			}
			table.SetCellValue("Offer Id", offer.Id)
			table.SetCellValue("Param Name", param.Name)
			table.SetCellValue("Param Value", param.Value)
			table.SetCellValue("Unit", param.Unit)
		}
	}
	return table
}

//structured object for simplier interpreting data for writing excel sheet.
//has list of columns with their nums
type Table struct {
	Columns map[string]int
	Rows    []Row
}

//substruct of table, contains cells. key of map is num if column
type Row struct {
	Cells map[int]string
}

//set cell value for current row.
func (t *Table) SetCellValue(column string, value string) {
	if !t.ContainsColumn(column) {
		t.AddColumn(column)
	}
	t.Rows[len(t.Rows)-1].Cells[t.Columns[column]] = value
}

//check has table given column yet or not
func (t *Table) ContainsColumn(column string) bool {
	_, ok := t.Columns[column]
	if ok {
		return true
	}
	return false
}

//add column to table
func (t *Table) AddColumn(column string) {
	if t.Columns == nil {
		t.Columns = make(map[string]int)
	}
	t.Columns[column] = len(t.Columns)
}

//add new row to table
func (t *Table) AddRow() error {
	if len(t.Rows) < 1048576 {
		row := Row{}
		row.Cells = make(map[int]string)
		t.Rows = append(t.Rows, row)
		return nil
	}
	return errors.New("Not wnough rows count.")
}

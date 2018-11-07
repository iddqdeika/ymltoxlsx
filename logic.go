package main

import "errors"

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
	Id   string `xml:"id,attr"`
	Name string `xml:",chardata"`
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
	Barcode       string        `xml:"barcode"`
	Params        []Param       `xml:"param"`
}

type Param struct {
	Name  string `xml:"name,attr"`
	Value string `xml:",chardata"`
}

//get table with offer's linear field
func (c *Yml_catalog) GetOfferTable() Table {
	table := Table{}
	for _, offer := range c.Shop.Offers.Offers {
		table.AddRow()
		table.SetCellValue("Id", offer.Id)
		table.SetCellValue("Available", offer.Available)
		table.SetCellValue("Url", offer.Url)
		table.SetCellValue("Price", offer.Price)
		table.SetCellValue("CurrencyID", offer.CurrencyId)
		table.SetCellValue("CategoryId", offer.CategoryId)
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
		table.SetCellValue("Barcode", offer.Barcode)

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

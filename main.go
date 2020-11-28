package main

import (
    "encoding/json"
    "fmt"
    "log"
    "os"
    "github.com/360EntSecGroup-Skylar/excelize"
    "github.com/urfave/cli/v2"
)

func main() {
    app := cli.NewApp()
    app.Name    = "excel2json"
    app.Usage   = "Convert Excel to JSON"
    app.Version = "0.0.1"
    app.Flags = []cli.Flag {
        &cli.StringFlag{
            Name: "file",
            Aliases: []string{"f"},
            Usage: "excel file path",
            Required: true,
        },
        &cli.StringSliceFlag{
            Name: "sheet",
            Aliases: []string{"s"},
            Usage: "parse excel sheet",
            Required: true,
        },
    }
    app.Action = func (c *cli.Context) error {
        excel, err := excelize.OpenFile(c.String("file"))
        if err != nil {
            fmt.Println(err)
            return err
        }
        var data = make(map[string]map[int]map[int]string)
        sheets := c.StringSlice("sheet")
        for _, sheet := range sheets {
            data[sheet] = make(map[int]map[int]string)
            rows := excel.GetRows(sheet)
            for yindex, row := range rows {
                for xindex, value := range row {
                    if value != "" {
                        if _, ok := data[sheet][yindex]; ! ok {
                            data[sheet][yindex] = make(map[int]string)
                        }
                        data[sheet][yindex][xindex] = value
                    }
                }
            }
        }
        if json, err := json.Marshal(data); err != nil {
            return err
        } else {
            fmt.Println(string(json))
            return nil
        }
    }
    err := app.Run(os.Args)
    if err != nil {
        log.Fatal(err)
    }
}

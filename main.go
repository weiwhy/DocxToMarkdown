package main

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"log"
	"os"
	"path/filepath"
	"strconv"
	"strings"
)

type Document struct {
	XMLName xml.Name `xml:"document"`
	Body    Body     `xml:"body"`
}

type Body struct {
	XMLName xml.Name `xml:"body"`
	Content []string
}

var ilvl = -1
var numId = -1
var bookmarkStart = -1
var sdt = false
var cols = 0
var rows = -1
var istbl = false
var dirPath string

type Relationships struct {
	XMLName      xml.Name       `xml:"Relationships"`
	Relationship []Relationship `xml:"Relationship"`
}

type Relationship struct {
	XMLName xml.Name `xml:"Relationship"`
	Id      string   `xml:"Id,attr"`
	Type    string   `xml:"Type,attr"`
	Target  string   `xml:"Target,attr"`
}

func main() {
	// 指定要提取的 Zip 文件
	if len(os.Args) < 2 {
		fmt.Println("DocxToMarkdown.exe 待转换格式文件.docx")
		os.Exit(0)
	}

	if _, err := os.Stat(os.Args[1]); err == nil {
		// 文件存在
		// 获取文件的绝对路径
		dirPath = filepath.Dir(os.Args[1])
	} else {
		fmt.Println("文件不存在", os.Args[1])
	}
	fileName := filepath.Base(os.Args[1])
	// 去掉文件后缀
	fileNameWithoutExt := fileName[:len(fileName)-len(filepath.Ext(fileName))]
	// 打开 Zip 文件进行读取
	reader, err := zip.OpenReader(os.Args[1])
	if err != nil {
		log.Fatal(err)
	}
	defer reader.Close()
	var xmlrelsData []byte
	var xmlxData []byte
	// 遍历 Zip 文件中的每个文件/目录
	for _, file := range reader.File {
		index := strings.Index(file.FileHeader.Name, "word/media/")
		if index == 0 {
			rc, err := file.Open()
			if err != nil {
				log.Fatal(err)
			}
			defer rc.Close()

			// 创建目标目录
			targetDir := dirPath + string(filepath.Separator) + "media"
			tmp := strings.ReplaceAll(file.FileHeader.Name, "word/", "")
			targetFile := dirPath + string(filepath.Separator) + tmp
			if err := os.MkdirAll(targetDir, os.ModePerm); err != nil {
				log.Fatal(err)
			}
			writer, err := os.Create(targetFile)
			if err != nil {
				log.Fatal(err)
			}
			defer writer.Close()

			// 将文件内容复制到目标文件
			_, err = io.Copy(writer, rc)
			if err != nil {
				log.Fatal(err)
			}
		} else if file.FileHeader.Name == "word/document.xml" {
			rc, err := file.Open()
			if err != nil {
				log.Fatal(err)
			}
			defer rc.Close()
			buf := new(bytes.Buffer)
			buf.ReadFrom(rc)
			xmlxData = buf.Bytes()
		} else if file.FileHeader.Name == "word/_rels/document.xml.rels" {
			rc, err := file.Open()
			if err != nil {
				log.Fatal(err)
			}
			defer rc.Close()
			buf := new(bytes.Buffer)
			buf.ReadFrom(rc)
			xmlrelsData = buf.Bytes()
		}
	}
	// 读取Word文档的document.xml.rels文件内容
	//xmlrelsData, err := ioutil.ReadFile("document.xml.rels.xml")
	//if err != nil {
	//	fmt.Println("Error reading file:", err)
	//	return
	//}
	var rels Relationships
	err = xml.Unmarshal(xmlrelsData, &rels)
	if err != nil {
		fmt.Println("Error parsing XML:", err)
		return
	}
	var id_image = make(map[string]string)
	for _, relationship := range rels.Relationship {
		id_image[relationship.Id] = relationship.Target
	}

	//解析document.xml
	//xmlxData, err := ioutil.ReadFile("表格.xml")
	//if err != nil {
	//	fmt.Println("Error reading file:", err)
	//	return
	//}
	xmlData := string(xmlxData)
	var doc Document
	err = xml.Unmarshal([]byte(xmlData), &doc)
	if err != nil {
		fmt.Println("Error parsing XML:", err)
		return
	}

	// 解析body标签下的内容，并按照顺序存储在Content字段中
	decoder := xml.NewDecoder(strings.NewReader(xmlData))
	var currentElement string
	//ilvl/numId/bookmarkStart(name)/sdt
	var content string
	for {
		token, err := decoder.Token()
		if err != nil {
			break
		}

		switch t := token.(type) {
		case xml.StartElement:

			currentElement = t.Name.Local
			if currentElement == "sdt" {
				sdt = true
			}
			if currentElement == "ilvl" {
				for _, attr := range t.Attr {
					if attr.Name.Local == "val" {
						tmp, _ := strconv.Atoi(attr.Value)
						ilvl = tmp
					}
				}
			}
			if currentElement == "numId" {
				for _, attr := range t.Attr {
					if attr.Name.Local == "val" {
						tmp, _ := strconv.Atoi(attr.Value)
						numId = tmp
					}
				}
			}
			if currentElement == "bookmarkStart" {
				for _, attr := range t.Attr {
					if attr.Name.Local == "name" {
						tmp, _ := strconv.Atoi(attr.Value)
						bookmarkStart = tmp
					}
				}
			}
			//blip
			if currentElement == "blip" {
				for _, attr := range t.Attr {
					if attr.Name.Local == "embed" {
						content = fmt.Sprintf("%s\n![%s](%s)\n", content, attr.Value, id_image[attr.Value])
					}
				}
			}
			if currentElement == "tbl" {
				istbl = true
				rows = 0
			}
			if currentElement == "tblGrid" {

			}
			if currentElement == "gridCol" {
				cols = cols + 1
			}
			if currentElement == "tr" {
				content = fmt.Sprintf("%s\n%s", content, "|")
				rows = rows + 1
			}
			if currentElement == "tc" {

			}
		case xml.CharData:
			//ilvl/numId/bookmarkStart(name)/sdt
			if currentElement == "t" {
				if sdt {

				} else if numId == 1 && ilvl == 0 { //1-0,numId,ilvl
					content = fmt.Sprintf("%s%s", content, fmt.Sprint("# ", strings.ReplaceAll(strings.ReplaceAll(string(t), "\t", ""), "\n", "")))
					ilvl = -1
					numId = -1
					bookmarkStart = -1
				} else if numId == 1 && ilvl == 1 { //1-1
					content = fmt.Sprintf("%s%s", content, fmt.Sprint("## ", strings.ReplaceAll(strings.ReplaceAll(string(t), "\t", ""), "\n", "")))
					ilvl = -1
					numId = -1
					bookmarkStart = -1
				} else if numId == 2 && ilvl == 0 { //2-0
					content = fmt.Sprintf("%s%s", content, fmt.Sprint("### ", strings.ReplaceAll(strings.ReplaceAll(string(t), "\t", ""), "\n", "")))
					ilvl = -1
					numId = -1
					bookmarkStart = -1
				} else {
					content = fmt.Sprintf("%s%s", content, fmt.Sprint(strings.ReplaceAll(strings.ReplaceAll(string(t), "\t", ""), "\n", "")))
				}

			}

		case xml.EndElement:
			if t.Name.Local == "p" {
				if istbl {
					content = fmt.Sprintf("%s%s", content, "")
				} else {
					content = fmt.Sprintf("%s%s", content, "\n")
				}
			}
			if t.Name.Local == "t" {
				currentElement = "x"
			}
			if t.Name.Local == "sdt" {
				sdt = false
			}
			if t.Name.Local == "tblGrid" {

			}
			if t.Name.Local == "tbl" {
				istbl = false
				rows = -1
			}
			if t.Name.Local == "tc" {
				content = fmt.Sprintf("%s%s", content, "|")
			}
			if t.Name.Local == "tr" {
				if rows == 1 {
					content = fmt.Sprintf("%s\n%s", content, "|")
					for i := 0; i < cols; i++ {
						content = fmt.Sprintf("%s%s", content, "----|")
					}
					cols = 0
				}
			}
		}
	}
	file, err := os.Create(dirPath + string(filepath.Separator) + fileNameWithoutExt + ".md")
	if err != nil {
		fmt.Println("创建MD文件错误:", err)
		return
	}
	defer file.Close()

	_, err = fmt.Fprintln(file, content)
	if err != nil {
		fmt.Println("向写MD文件写内容错误:", err)
		return
	}

}

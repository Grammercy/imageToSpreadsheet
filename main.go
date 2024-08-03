package main

import (
  "fmt"
  "os"
  "strconv"
  "image"
  // "time"
  _ "image/png"
  _ "image/jpeg"
  _ "image/gif"
  "image/color"
  "golang.org/x/image/bmp"
  "golang.org/x/image/tiff"
  "golang.org/x/image/webp"
  "github.com/tealeg/xlsx"
  "github.com/nfnt/resize"
)

type Pixel struct {
  R int
  G int
  B int
  A int
}

func main() {
  // start := time.Now()
  img, width := processArgs(os.Args)
  file := xlsx.NewFile()

  createExcelSheet(img, "image", width, file)
  
  err := file.Save("output.xlsx")
  
  if err != nil {
    fmt.Println(err)
  }
  fmt.Println("Written successfully.")
  // fmt.Println("Time taken ", time.Since(start))
}

func processArgs(args []string) ([][]Pixel, float64) {
  var num int
  width := 2.0
  var img [][]Pixel
  switch len(args) {
  case 2:
    img = imgToPixelArr(getAndProcessImage(os.Args[1]))
  case 3: 
    num, _ = strconv.Atoi(os.Args[2])
    img = imgToPixelArr(resize.Resize(uint(num), 0, getAndProcessImage(os.Args[1]), resize.Lanczos3))
  case 4:
    num, _ = strconv.Atoi(os.Args[2])
    num2, _ := strconv.Atoi(os.Args[3])
    img = imgToPixelArr(resize.Resize(uint(num), uint(num2), getAndProcessImage(os.Args[1]), resize.Lanczos3))
  case 5: 
    num, _ = strconv.Atoi(os.Args[2])
    num2, _ := strconv.Atoi(os.Args[3])
    img = imgToPixelArr(resize.Resize(uint(num), uint(num2), getAndProcessImage(os.Args[1]), resize.Lanczos3))
    width, _ = strconv.ParseFloat(os.Args[4], 64)
  default:
    fmt.Println("Invalid number of arguments, all arguments after 1 are optional\n usage: imageToSpreadsheet localPathToImage<only takes png>` width<optional, preserves aspect ratio> height<optional, does not preserve aspect ratio> widthOfCell<optional>")
    os.Exit(128)
  }
  return img, width

}

func createExcelSheet(img [][]Pixel, sheetName string, width float64, file *xlsx.File) {
  sheet, err := file.AddSheet(sheetName)
  
  if err != nil {
    fmt.Println(err)
    return
  }
  
  sheet.SetColWidth(0, len(img[0]), width)

  
  for y := 0; y < len(img); y++ {
    row := sheet.AddRow()
    row.SetHeight(7 * width)
    for x := 0; x < len(img[y]); x++ {
      fillCell(x, y, img, row.AddCell())
    }
  }

}

func fillCell(x, y int, img [][]Pixel, cell *xlsx.Cell) {
  style := xlsx.NewStyle()
  // fmt.Println("Filling (", x, y, ") with ", pixelToHex(img[y][x]))
  style.Fill = *xlsx.NewFill("solid", pixelToHex(img[y][x]), pixelToHex(img[y][x]))
  style.ApplyFill = true
  cell.SetStyle(style)

}

func pixelToHex(pixel Pixel) string {
  return fmt.Sprintf("%02X%02X%02X", pixel.R, pixel.G, pixel.B)
}

func imgToPixelArr(img image.Image) [][]Pixel {
  width := img.Bounds().Max.X 
  height := img.Bounds().Max.Y

  var arr [][]Pixel

  for y := 0; y < height; y++ {
    var row []Pixel
    for x := 0; x < width; x++ {
      row = append(row, PixelToPixel(img.At(x, y)))
    }
    arr = append(arr, row)
  }
  return arr
}

func getAndProcessImage(path string) image.Image {
  imgPath := "./" + path
  imgFile, err := os.Open(imgPath)
  
  if err != nil {
    fmt.Println("Image path invalid")
    os.Exit(1)
  }
  
  defer imgFile.Close()
  
  img, _, err := image.Decode(imgFile)
  
  if err == nil {
    return img
  }
  
  img, err = webp.Decode(imgFile)
  
  if err == nil {
    return img
  }

  img, err = bmp.Decode(imgFile)

  if err == nil {
    return img
  }

  img, err = tiff.Decode(imgFile)

  if err == nil {
    return img
  }
  
  fmt.Println("Error decoding image, unsupported format?")
  os.Exit(1)
  return nil
}

func PixelToPixel(pixel color.Color) Pixel {
  r, g, b, a := pixel.RGBA()
  return Pixel{int(r / 257), int(g / 257), int(b / 257), int(a / 257)}
}

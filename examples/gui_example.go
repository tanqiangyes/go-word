package main

import (
	"fmt"
	"log"

	"github.com/tanqiangyes/go-word/pkg/gui"
)

// main 主函数 - GUI示例
func main() {
	fmt.Println("启动 Go Word GUI 应用...")
	
	// 创建GUI实例
	app := gui.NewGUI()
	if app == nil {
		log.Fatal("GUI应用创建失败")
	}
	
	fmt.Println("GUI应用已创建，正在启动...")
	
	// 运行GUI应用
	app.Run()
	
	fmt.Println("GUI应用已关闭")
}

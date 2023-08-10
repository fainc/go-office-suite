package value

type Sheet struct {
	SheetName string      `dc:"表名称，不设置默认顺序Sheet1,2,3"`
	Desc      []Desc      `dc:"表内头部文档描述"`
	Field     []Field     `dc:"数据键"`
	Rows      interface{} `dc:"数据集"`
}

type Desc struct {
	Text     string  `dc:"文档描述文字"`
	Column   int     `dc:"指定竖向合并行数(横向合并列根据数据列自动计算)"`
	Align    string  `dc:"文字水平对齐方式"`
	FontSize float64 `dc:"字体大小"`
}

type Field struct {
	Name        string  `dc:"数据键名称"`
	Index       string  `dc:"数据索引"`
	Child       []Field `dc:"子键"`
	RenderImage bool    `dc:"是否将当前列以图片渲染（网络图需要下载，可能生成较慢或失败）"`
}

type File struct {
	SavePath    string `dc:"文件保存路径"`
	ActiveSheet int    `dc:"设置默认表，默认0"`
}

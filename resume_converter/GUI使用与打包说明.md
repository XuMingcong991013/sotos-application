# GUI使用与打包说明

## 源码运行

确保项目根目录存在已配置的 `.env`，然后执行：

```powershell
python gui.py
```

点击“选择文件”可选择单个简历，也可按住 Ctrl/Shift 一次选择多个简历。
选择后，输出文件夹会自动建议为首个文件所在目录下的
`标准简历输出`，仍可自行更改。GUI只处理用户明确选中的文件。

## 构建Windows EXE

先在用于构建的Python环境安装PyInstaller：

```powershell
python -m pip install PyInstaller
```

然后在项目根目录运行：

```powershell
powershell -ExecutionPolicy Bypass -File .\build_exe.ps1
```

构建结果：

```text
dist/简历标准化处理工具/
  简历标准化处理工具.exe
  _internal/
```

这是目录版程序，必须整体分发，不能只复制EXE。当前属于内部测试版本，
构建时会把项目根目录的`.env`放入PyInstaller资源；用户不需要在EXE旁
另放`.env`。

注意：PyInstaller资源可以被解包，这种做法只适合当前内部测试，不能作为
正式版本的密钥安全方案。正式分发前应改为内网后端代管密钥或Windows凭据保护。

第一版使用`onedir + noconsole`，启动速度和依赖加载稳定性优先。程序模板
已经作为只读资源打包，用户不需要单独复制`templates`目录。

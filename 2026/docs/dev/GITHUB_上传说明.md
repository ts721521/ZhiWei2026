# 用 GitHub 工具上传代码与文档

## 一、已为你安装的工具

已通过 **winget** 安装：

- **Git for Windows**（版本约 2.53）  
- **GitHub CLI（gh）**（版本约 2.87）

安装完成后，请**重新打开一个终端**（或重启 Cursor），这样 `git` 和 `gh` 才会在 PATH 中生效。

---

## 二、首次使用：登录 GitHub

在**新终端**中执行：

```powershell
gh auth login
```

按提示选择：

- **GitHub.com**
- 协议选 **HTTPS** 或 **SSH**（任选其一）
- 认证方式选 **Login with a web browser**，按提示在浏览器中完成登录

登录成功后，即可用下面的方式上传代码。

---

## 三、上传当前代码与文档

### 方式 1：用脚本（推荐）

在仓库根目录 `d:\GitHub\ZhiWei2026` 下双击或运行：

```bat
upload_to_github.bat
```

如需自定义提交信息：

```bat
upload_to_github.bat "你的提交说明"
```

### 方式 2：手动命令

在仓库根目录打开终端，执行：

```powershell
git add -A
git commit -m "Sync: code and docs"
git push -u origin main
```

---

## 四、当前仓库信息

- **远程地址**：https://github.com/ts721521/ZhiWei2026.git  
- **默认分支**：`main`  

`.gitignore` 已配置好，不会把 `config.json`、`logs/`、`tasks/`、`__pycache__` 等敏感或临时文件提交上去。

---

## 五、若终端里仍找不到 git / gh

若重新打开终端后仍提示找不到 `git` 或 `gh`，可手动把安装路径加入当前会话的 PATH：

```powershell
$env:Path += ";C:\Program Files\Git\cmd"
# GitHub CLI 通常在：
$env:Path += ";C:\Program Files\GitHub CLI"
```

然后再执行 `git --version` 和 `gh --version` 检查。

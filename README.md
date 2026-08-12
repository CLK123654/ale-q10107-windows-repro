# API月度计费

本仓库保存API平台月度计费任务。artifacts目录包含输入数据、标准交付和两份工作簿，task目录包含业务说明，verification目录保存Windows检查程序。

工作流使用windows-2025和Node.js24内置SQLite。它会从输入包运行任务，核对修复数据库与三份业务报表，并在完成后上传运行证据。

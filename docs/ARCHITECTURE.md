# Architecture / 架构

The application follows a single production UI and a layered processing pipeline.

```text
PySide6 UI
  ├─ SplitPage -> SplitTaskConfig
  └─ MergePage -> MergeTaskConfig / AdvancedMergeTaskConfig
                         │
                         ▼
               ProcessingWorker (QThread)
                         │
       scan -> inspect -> read -> transform -> write
                         │
                         ▼
            reconciliation + reports + history
```

- `config/`: centralized Chinese/English text, settings, and constants.
- `models/`: immutable task choices and execution/result data.
- `services/`: workbook inspection, readers, split/merge engines, writers, and reports.
- `workers/`: background `QObject` workers; workbook I/O does not run in button handlers.
- `ui/`: the only production PySide6 interface.

The default output policy is automatic renaming. Source paths are never used as output targets by
the UI, and every field-based split is reconciled against the number of input rows.

应用只有一套正式 UI。界面负责生成任务配置，后台线程负责工作簿读写，服务层执行拆分、合并、
输出与行数核对。默认同名策略为自动编号，不覆盖源文件或已有输出。

## Word: Apply a highlight to all tracked changes
在提交返修稿时， ISME J 要求我们提交一个正文新旧版本的比较结果，并且高亮差异部分。

## 方法
### 1. 新旧版本比较
打开word -> 审阅(review) -> 比较(compare) ->

![example.png](https://s1.ax1x.com/2020/06/16/NFunUI.md.png)


### 2. 高亮差异部分

点击 开发工具(developer)-> Visual Basic -> 右键点击project -> 插入(insert) -> 模块(module)
再粘贴下面的代码:
```
Sub tracked_to_highlighted()           
    tempState = ActiveDocument.TrackRevisions
    ActiveDocument.TrackRevisions = Flase    
    For Each Change In ActiveDocument.Revisions
        Set myRange = Change.Range
        myRange.HighlightColorIndex = wdYellow           
    Next    
    ActiveDocument.TrackRevisions = tempState
End Sub
```

你将看到
[![NFQVrF.png](https://s1.ax1x.com/2020/06/16/NFQVrF.png)](https://imgchr.com/i/NFQVrF)

然后运行一下这个模块即可. **ps**截图代码有误，VB脚本请用上面可复制的


#
PS1. 该方法对于bibliography部分会报错
>Run time error '5825'.
Objects have been deleted.

可以先忽略bibliography部分的高亮，重新添加一次reference即可


PS2. word文档比较还有一个bug是：某些本来只是部分被修改但并未被完全删除的段落会被识别为整段删除后新加一段，这种情况暂时没找到解决方案，需手动处理T-T

### Ref:
https://cybertext.wordpress.com/2018/11/22/word-apply-a-highlight-to-all-tracked-changes/

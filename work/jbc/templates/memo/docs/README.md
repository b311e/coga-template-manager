# JBC Memo Template Documentation

Create the jbcLetterhead template from the jbcLetterhead template.

1) Create a copy of the jbcLetterhead
2) In the body of the document, press `Ctrl + F9` to insert fields
3) Type `AUTOTEXT Memo` inside the brackets

OR

Just paste in the `document.xml`:

```xml
        <w:p>
            <w:r>
                <w:fldChar w:fldCharType="begin"/>
            </w:r>
            <w:r>
                <w:instrText xml:space="preserve"> AUTOTEXT Memo </w:instrText>
            </w:r>
            <w:r>
                <w:fldChar w:fldCharType="end"/>
            </w:r>
            <w:bookmarkStart w:id="0" w:name="_GoBack"/>
            <w:bookmarkEnd w:id="0"/>
        </w:p>
```



This will pull the AutoText called "Memo" into the document, so be sure that is actually included in the Letterhead Template.
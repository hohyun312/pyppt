# pyppt
파이썬으로 `.PPTX`를 조작할 수 있게 해주는 [python-pptx](https://python-pptx.readthedocs.io/en/latest/)를 조금 더 쉽게 사용할 수 있게 해줍니다.

# Overview
```python
import pyppt

ppt = pyppt()
ppt.add_slide()
ppt.set_slide_size(30, 10)
ppt.add_textbox(1, left=3, top=3, width=10, height=5, name='textbox1')
ppt.edit_text(1, 'textbox1', '테스트입니다1', bold=True)
ppt.edit_text(1, 'textbox1', '테스트입니다2', italic=True, color='red', alignment='left', clear=False)
ppt.edit_text(1, 'textbox1', '테스트입니다3', font='궁서체', color='green', clear=False)
ppt.save('test.pptx')
```

![테스트 이미지](./images/test_img.jpg)

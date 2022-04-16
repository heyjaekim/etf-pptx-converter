from pptx import Presentation


    
def chart_to_picture(slide, placeholder_number, fig_path):
    picture_placeholder = slide.shapes[placeholder_number]
    picture_placeholder.insert_picture(fig_path)

def chart_to_picture_style(slide, placeholder_number, fig_path):
    slide.placeholders[placeholder_number].insert_picture(fig_path)


save_path = r'C:/ETF'

prs = Presentation(save_path + '/etf_kd_ver1.pptx')
for idx, slide_layout in enumerate(prs.slide_layouts):
    slide = prs.slides.add_slide(slide_layout)
    for shape in slide.placeholders:
        print(idx, shape.placeholder_format.idx, shape.name, shape.top, shape.left)

prs = Presentation(save_path + '/etf_kd_ver1.pptx')    # open a pptx file
layout = prs.slide_layouts[1]
slide = prs.slides.add_slide(layout)
mult_fig_path = save_path + '/mult_fig.png'
chart_to_picture(slide, 17, mult_fig_path)
prs.save(save_path + '/etf_kd_ver1.pptx')
chart_to_picture(slide, 18, mult_fig_path)
prs.save(save_path + '/etf_kd_ver1.pptx')
chart_to_picture(slide, 19, mult_fig_path)
prs.save(save_path + '/etf_kd_ver1.pptx')
chart_to_picture(slide, 20, mult_fig_path)
prs.save(save_path + '/etf_kd_ver1.pptx')
chart_to_picture(slide, 21, mult_fig_path)
prs.save(save_path + '/etf_kd_ver1.pptx')
chart_to_picture(slide, 22, mult_fig_path)
prs.save(save_path + '/etf_kd_ver1.pptx')
chart_to_picture(slide, 23, mult_fig_path)
prs.save(save_path + '/etf_kd_ver1.pptx')
chart_to_picture(slide, 24, mult_fig_path)
prs.save(save_path + '/etf_kd_ver1.pptx')
chart_to_picture(slide, 25, mult_fig_path)
prs.save(save_path + '/etf_kd_ver1.pptx')
chart_to_picture(slide, 26, mult_fig_path)
prs.save(save_path + '/etf_kd_ver1.pptx')
chart_to_picture(slide, 27, mult_fig_path)
prs.save(save_path + '/etf_kd_ver1.pptx')
chart_to_picture(slide, 28, mult_fig_path)
prs.save(save_path + '/etf_kd_ver1.pptx')
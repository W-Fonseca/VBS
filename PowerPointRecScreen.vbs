ola

Sub ola()

Set objPPT = CreateObject("PowerPoint.Application")
objPPT.Visible = True
Set objPresentation = objPPT.Presentations.Add
Set objSlide = objPresentation.Slides.Add(1, 2)
objSlide.Application.CommandBars.ExecuteMso("ObjectScreenRecording")


end sub

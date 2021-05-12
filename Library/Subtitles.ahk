; Script:    Vis2.ahk
; Author:    iseahound
; Date:      2017-08-19
; Recent:    2018-04-04

;#include <Gdip_All>
;#include <JSON>


; ImageIdentify() - Label and identify objects in images.
ImageIdentify(image:="", search:="", options:=""){
   return Vis2.ImageIdentify(image, search, options)
}

; OCR() - Convert pictures of text into text.
OCR(image:="", language:="", options:=""){
   return Vis2.OCR(image, language, options)
}

class Vis2 {
   class Graphics {

      static pToken, Gdip := 0

      Startup(){
         global pToken
         return Vis2.Graphics.pToken := (Vis2.Graphics.Gdip++ > 0) ? Vis2.Graphics.pToken : (pToken) ? pToken : Gdip_Startup()
      }

      Shutdown(){
         global pToken
         return Vis2.Graphics.pToken := (--Vis2.Graphics.Gdip <= 0) ? ((pToken) ? pToken : Gdip_Shutdown(Vis2.Graphics.pToken)) : Vis2.Graphics.pToken
      }

      Name(){
         VarSetCapacity(UUID, 16, 0)
         if (DllCall("rpcrt4.dll\UuidCreate", "ptr", &UUID) != 0)
             return (ErrorLevel := 1) & 0
         if (DllCall("rpcrt4.dll\UuidToString", "ptr", &UUID, "uint*", suuid) != 0)
             return (ErrorLevel := 2) & 0
         return A_TickCount "n" SubStr(StrGet(suuid), 1, 8), DllCall("rpcrt4.dll\RpcStringFree", "uint*", suuid)
      }

      class Area{

         ScreenWidth := A_ScreenWidth, ScreenHeight := A_ScreenHeight,
         action := ["base"], x := [0], y := [0], w := [1], h := [1], a := ["top left"], q := ["bottom right"]

         __New(name := "", color := "0x7FDDDDDD") {
            this.name := name := (name == "") ? Vis2.Graphics.Name() "_Graphics_Area" : name "_Graphics_Area"
            this.color := color

            Vis2.Graphics.Startup()
            Gui, %name%:New, +LastFound +AlwaysOnTop -Caption -DPIScale +E0x80000 +ToolWindow +hwndSecretName, % this.name
            Gui, %name%:Show, % (this.isDrawable()) ? "NoActivate" : ""
            this.hwnd := SecretName
            this.hbm := CreateDIBSection(this.ScreenWidth, this.ScreenHeight)
            this.hdc := CreateCompatibleDC()
            this.obm := SelectObject(this.hdc, this.hbm)
            this.G := Gdip_GraphicsFromHDC(this.hdc)
            Gdip_SetSmoothingMode(this.G, 4) ;Adds one clickable pixel to the edge.
            this.pBrush := Gdip_BrushCreateSolid(this.color)
         }

         __Delete(){
            Vis2.Graphics.Shutdown()
         }

         Destroy(){
            Gdip_DeleteBrush(this.pBrush)
            SelectObject(this.hdc, this.obm)
            DeleteObject(this.hbm)
            DeleteDC(this.hdc)
            Gdip_DeleteGraphics(this.G)
            Gui, % this.name ":Destroy"
         }

         Hide(){
            DllCall("ShowWindow", "ptr",this.hWnd, "int",0)
         }

         Show(){ ; NoActivate
            DllCall("ShowWindow", "ptr",this.hWnd, "int",8)
         }

         ToggleVisible(){
            this.isVisible() ? this.Hide() : this.Show()
         }

         isVisible(){
            return DllCall("IsWindowVisible", "ptr",this.hWnd)
         }

         isDrawable(win := "A"){
             static WM_KEYDOWN := 0x100,
             static WM_KEYUP := 0x101,
             static vk_to_use := 7
             ; Test whether we can send keystrokes to this window.
             ; Use a virtual keycode which is unlikely to do anything:
             PostMessage, WM_KEYDOWN, vk_to_use, 0,, % win
             if !ErrorLevel
             {   ; Seems best to post key-up, in case the window is keeping track.
                 PostMessage, WM_KEYUP, vk_to_use, 0xC0000000,, % win
                 return true
             }
             return false
         }

         DetectScreenResolutionChange(){
            if (this.ScreenWidth != A_ScreenWidth || this.ScreenHeight != A_ScreenHeight) {
               this.ScreenWidth := A_ScreenWidth, this.ScreenHeight := A_ScreenHeight
               SelectObject(this.hdc, this.obm)
               DeleteObject(this.hbm)
               DeleteDC(this.hdc)
               Gdip_DeleteGraphics(this.G)
               this.hbm := CreateDIBSection(this.ScreenWidth, this.ScreenHeight)
               this.hdc := CreateCompatibleDC()
               this.obm := SelectObject(this.hdc, this.hbm)
               this.G := Gdip_GraphicsFromHDC(this.hdc)
               Gdip_SetSmoothingMode(this.G, 4)
            }
         }

         Redraw(x, y, w, h){
            Critical On
            this.DetectScreenResolutionChange()
            Gdip_GraphicsClear(this.G)
            Gdip_FillRectangle(this.G, this.pBrush, x, y, w, h)
            if (this.coordinates)
               Vis2.Graphics.Subtitle.Draw("x: " x " │ y: " y " │ w: " w " │ h: " h
                  , {"a":"top_right", "x":"right", "y":"top", "color":"Black"}, {"font":"Lucida Sans Typewriter", "size":"1.67%"}, this.G)
            UpdateLayeredWindow(this.hwnd, this.hdc, 0, 0, this.ScreenWidth, this.ScreenHeight)
            Critical Off
         }

         ChangeColor(color){
            this.color := color
            Gdip_DeleteBrush(this.pBrush)
            this.pBrush := Gdip_BrushCreateSolid(this.color)
            this.Redraw(this.x[this.x.MaxIndex()], this.y[this.y.MaxIndex()], this.w[this.w.MaxIndex()], this.h[this.h.MaxIndex()])
         }

         ShowCoordinates(){
            this.coordinates := true
            this.Redraw(this.x[this.x.MaxIndex()], this.y[this.y.MaxIndex()], this.w[this.w.MaxIndex()], this.h[this.h.MaxIndex()])
         }

         HideCoordinates(){
            this.coodinates := false
            this.Redraw(this.x[this.x.MaxIndex()], this.y[this.y.MaxIndex()], this.w[this.w.MaxIndex()], this.h[this.h.MaxIndex()])
         }

         ToggleCoordinates(){
            this.coordinates := !this.coordinates
            this.Redraw(this.x[this.x.MaxIndex()], this.y[this.y.MaxIndex()], this.w[this.w.MaxIndex()], this.h[this.h.MaxIndex()])
         }

         Propagate(v){
            this.a[v] := (this.a[v] == "") ? this.a[v-1] : this.a[v]
            this.q[v] := (this.q[v] == "") ? this.q[v-1] : this.q[v]
            this.x[v] := (this.x[v] == "") ? this.x[v-1] : this.x[v]
            this.y[v] := (this.y[v] == "") ? this.y[v-1] : this.y[v]
            this.w[v] := (this.w[v] == "") ? this.w[v-1] : this.w[v]
            this.h[v] := (this.h[v] == "") ? this.h[v-1] : this.h[v]
         }

         BackPropagate(pasts){
            action := this.action.pop()
            a := this.a.pop()
            q := this.q.pop()
            x := this.x.pop()
            y := this.y.pop()
            w := this.w.pop()
            h := this.h.pop()

            dx := x - this.x[pasts-1]
            dy := y - this.y[pasts-1]
            dw := w - this.w[pasts-1]
            dh := h - this.h[pasts-1]

            i := pasts-1
            while (i >= 1) {
               this.x[i] += dx
               this.y[i] += dy
               this.w[i] += dw
               this.h[i] += dh
               i--
            }
         }

         Converge(v := ""){
            v := (v) ? v : this.action.MaxIndex()

            if (v > 2) {
               this.action := [this.action[v-1], this.action[v]]
               this.a := [this.a[v-1], this.a[v]]
               this.q := [this.q[v-1], this.q[v]]
               this.x := [this.x[v-1], this.x[v]]
               this.y := [this.y[v-1], this.y[v]]
               this.w := [this.w[v-1], this.w[v]]
               this.h := [this.h[v-1], this.h[v]]
            }
         }

         Debug(function){
            v := (v) ? v : this.action.MaxIndex()
            Tooltip % function "`t" v . "`n" v-1 ": " this.action[v-1]
               . "`n" this.x[v-2] ", " this.y[v-2] ", " this.w[v-2] ", " this.h[v-2]
               . "`n" this.x[v-1] ", " this.y[v-1] ", " this.w[v-1] ", " this.h[v-1]
               . "`n" this.x[v] ", " this.y[v] ", " this.w[v] ", " this.h[v]
               . "`nAnchor:`t" this.a[v] "`nMouse:`t" this.q[v] "`t" this.isMouseInside()
         }

         Hover(){
            CoordMode, Mouse, Screen
            MouseGetPos, x_hover, y_hover
            this.x_hover := x_hover
            this.y_hover := y_hover

            ; Resets the stack to 1.
            if (A_ThisFunc != this.action[this.action.MaxIndex()]){
               this.action := [A_ThisFunc]
               this.a := [this.a.pop()]
               this.q := [this.q.pop()]
               this.x := [this.x.pop()]
               this.y := [this.y.pop()]
               this.w := [this.w.pop()]
               this.h := [this.h.pop()]
            }
         }

         Origin(v := ""){
            CoordMode, Mouse, Screen
            MouseGetPos, x_mouse, y_mouse

            if (A_ThisFunc != this.action[this.action.MaxIndex()]){
               this.action.push(A_ThisFunc)
               this.x_hover := x_mouse
               this.y_hover := y_mouse
            }

            v := (v) ? v : this.action.MaxIndex()

            if (x_mouse != this.x_last || y_mouse != this.y_last) {
               this.x_last := x_mouse, this.y_last := y_mouse

               this.x[v] := x_mouse
               this.y[v] := y_mouse

               this.Propagate(v)
               this.Redraw(x_mouse, y_mouse, 1, 1) ;stabilize x/y corrdinates in window spy.
            }
         }

         Draw(v := ""){
            CoordMode, Mouse, Screen
            MouseGetPos, x_mouse, y_mouse

            if (A_ThisFunc == this.action[this.action.MaxIndex()-1]){
               this.BackPropagate(this.action.MaxIndex())
               this.x_hover := x_mouse
               this.y_hover := y_mouse
               pass := 1
            }
            if (A_ThisFunc != this.action[this.action.MaxIndex()]){
               this.Converge()
               this.action.push(A_ThisFunc)
               this.x_hover := x_mouse
               this.y_hover := y_mouse
               pass := 1
            }

            v := (v) ? v : this.action.MaxIndex()
            dx := x_mouse - this.x_hover
            dy := y_mouse - this.y_hover
            xr := (x_mouse > this.x[v-1]) ? 1 : 0
            yr := (y_mouse > this.y[v-1]) ? 1 : 0

            if (pass == 1 || x_mouse != this.x_last || y_mouse != this.y_last) {
               this.x_last := x_mouse, this.y_last := y_mouse

               this.x[v] := (xr) ? this.x[v-1] : x_mouse
               this.y[v] := (yr) ? this.y[v-1] : y_mouse
               this.w[v] := (xr) ? x_mouse - this.x[v-1] : this.x[v-1] - x_mouse
               this.h[v] := (yr) ? y_mouse - this.y[v-1] : this.y[v-1] - y_mouse

               this.a[v] := (xr && yr) ? "top left" : (xr && !yr) ? "bottom left" : (!xr && yr) ? "top right" : "bottom right"
               this.q[v] := (xr && yr) ? "bottom right" : (xr && !yr) ? "top right" : (!xr && yr) ? "bottom left" : "top left"

               this.Propagate(v)
               this.Redraw(this.x[v], this.y[v], this.w[v], this.h[v])
            }
         }

         Move(v := ""){
            CoordMode, Mouse, Screen
            MouseGetPos, x_mouse, y_mouse

            if (A_ThisFunc != this.action[this.action.MaxIndex()]){
               this.Converge()
               this.action.push(A_ThisFunc)
               this.x_hover := x_mouse
               this.y_hover := y_mouse
               pass := 1
            }

            v := (v) ? v : this.action.MaxIndex()
            dx := x_mouse - this.x_hover
            dy := y_mouse - this.y_hover

            if (pass == 1 || x_mouse != this.x_last || y_mouse != this.y_last) {
               this.x_last := x_mouse, this.y_last := y_mouse

               this.x[v] := this.x[v-1] + dx
               this.y[v] := this.y[v-1] + dy

               this.Propagate(v)
               this.Redraw(this.x[v], this.y[v], this.w[v], this.h[v])
            }
         }

         ResizeCorners(v := ""){
            CoordMode, Mouse, Screen
            MouseGetPos, x_mouse, y_mouse

            if (A_ThisFunc != this.action[this.action.MaxIndex()]){
               this.Converge()
               this.action.push(A_ThisFunc)
               this.x_hover := x_mouse
               this.y_hover := y_mouse
               pass := 1
            }

            v := (v) ? v : this.action.MaxIndex()
            xr := this.x_hover - this.x[v-1] - (this.w[v-1] / 2)
            yr := this.y[v-1] - this.y_hover + (this.h[v-1] / 2) ; Keep Change Change
            dx := x_mouse - this.x_hover
            dy := y_mouse - this.y_hover

            if (pass == 1 || x_mouse != this.x_last || y_mouse != this.y_last) {
               this.x_last := x_mouse, this.y_last := y_mouse

               if (xr < -1 && yr > 1) {
                  r := "top left"
                  this.x[v] := this.x[v-1] + dx
                  this.y[v] := this.y[v-1] + dy
                  this.w[v] := this.w[v-1] - dx
                  this.h[v] := this.h[v-1] - dy
               }
               if (xr >= -1 && yr > 1) {
                  r := "top right"
                  this.x[v] := this.x[v-1]
                  this.y[v] := this.y[v-1] + dy
                  this.w[v] := this.w[v-1] + dx
                  this.h[v] := this.h[v-1] - dy
               }
               if (xr < -1 && yr <= 1) {
                  r := "bottom left"
                  this.x[v] := this.x[v-1] + dx
                  this.y[v] := this.y[v-1]
                  this.w[v] := this.w[v-1] - dx
                  this.h[v] := this.h[v-1] + dy
               }
               if (xr >= -1 && yr <= 1) {
                  r := "bottom right"
                  this.x[v] := this.x[v-1]
                  this.y[v] := this.y[v-1]
                  this.w[v] := this.w[v-1] + dx
                  this.h[v] := this.h[v-1] + dy
               }

               this.Propagate(v)
               this.Redraw(this.x[v], this.y[v], this.w[v], this.h[v])
            }
         }

         ; This works by finding the line equations of the diagonals of the rectangle.
         ; To identify the quadrant the cursor is located in, the while loop compares it's y value
         ; with the function line values f(x) = m * xr and y = -m * xr.
         ; So if yr is below both theoretical y values, then we know it's in the bottom quadrant.
         ; Be careful with this code, it converts the y plane inversely to match the Decartes tradition.

         ; Safety features include checking for past values to prevent flickering
         ; Sleep statements are required in every while loop.

         ResizeEdges(v := ""){
            CoordMode, Mouse, Screen
            MouseGetPos, x_mouse, y_mouse

            if (A_ThisFunc != this.action[this.action.MaxIndex()]){
               this.Converge()
               this.action.push(A_ThisFunc)
               this.x_hover := x_mouse
               this.y_hover := y_mouse
               pass := 1
            }

            v := (v) ? v : this.action.MaxIndex()
            m := -(this.h[v-1] / this.w[v-1])                              ; slope (dy/dx)
            xr := this.x_hover - this.x[v-1] - (this.w[v-1] / 2)           ; draw a line across the center
            yr := this.y[v-1] - this.y_hover + (this.h[v-1] / 2)           ; draw a vertical line halfing it
            dx := x_mouse - this.x_hover
            dy := y_mouse - this.y_hover

            if (pass == 1 || x_mouse != this.x_last || y_mouse != this.y_last) {
               this.x_last := x_mouse, this.y_last := y_mouse

               if (m * xr >= yr && yr > -m * xr)
                  r := "left",                this.x[v] := this.x[v-1] + dx,               this.w[v] := this.w[v-1] - dx
               if (m * xr < yr && yr > -m * xr)
                  r := "top",                 this.y[v] := this.y[v-1] + dy,               this.h[v] := this.h[v-1] - dy
               if (m * xr < yr && yr <= -m * xr)
                  r := "right",   this.w[v] := this.w[v-1] + dx
               if (m * xr >= yr && yr <= -m * xr)
                  r := "bottom",  this.h[v] := this.h[v-1] + dy

               this.Propagate(v)
               this.Redraw(this.x[v], this.y[v], this.w[v], this.h[v])
            }
         }

         isMouseInside(){
            CoordMode, Mouse, Screen
            MouseGetPos, x_mouse, y_mouse
            return (x_mouse >= this.x[this.x.MaxIndex()]
               && x_mouse <= this.x[this.x.MaxIndex()] + this.w[this.w.MaxIndex()]
               && y_mouse >= this.y[this.y.MaxIndex()]
               && y_mouse <= this.y[this.y.MaxIndex()] + this.h[this.h.MaxIndex()])
         }

         isMouseOutside(){
            return !this.isMouseInside()
         }

         isMouseOnCorner(){
            CoordMode, Mouse, Screen
            MouseGetPos, x_mouse, y_mouse
            return (x_mouse == this.x[this.x.MaxIndex()] || x_mouse == this.x[this.x.MaxIndex()] + this.w[this.w.MaxIndex()])
               && (y_mouse == this.y[this.y.MaxIndex()] || y_mouse == this.y[this.y.MaxIndex()] + this.h[this.h.MaxIndex()])
         }

         isMouseOnEdge(){
            CoordMode, Mouse, Screen
            MouseGetPos, x_mouse, y_mouse
            return ((x_mouse >= this.x[this.x.MaxIndex()] && x_mouse <= this.x[this.x.MaxIndex()] + this.w[this.w.MaxIndex()])
               && (y_mouse == this.y[this.y.MaxIndex()] || y_mouse == this.y[this.y.MaxIndex()] + this.h[this.h.MaxIndex()]))
               OR ((y_mouse >= this.y[this.y.MaxIndex()] && y_mouse <= this.y[this.y.MaxIndex()] + this.h[this.h.MaxIndex()])
               && (x_mouse == this.x[this.x.MaxIndex()] || x_mouse == this.x[this.x.MaxIndex()] + this.w[this.w.MaxIndex()]))
         }

         isMouseStopped(){
            CoordMode, Mouse, Screen
            MouseGetPos, x_mouse, y_mouse
            return x_mouse == this.x_last && y_mouse == this.y_last
         }

         ScreenshotRectangle(){
            x := this.x1(), y := this.y1(), w := this.width(), h := this.height()
            return (w > 0 && h > 0) ? (x "|" y "|" w "|" h) : ""
         }

         x1(){
            return this.x[this.x.MaxIndex()]
         }

         x2(){
            return this.x[this.x.MaxIndex()] + this.w[this.w.MaxIndex()]
         }

         y1(){
            return this.y[this.y.MaxIndex()]
         }

         y2(){
            return this.y[this.y.MaxIndex()] + this.h[this.h.MaxIndex()]
         }

         width(){
            return this.w[this.w.MaxIndex()]
         }

         height(){
            return this.h[this.h.MaxIndex()]
         }
      }

      Class CustomFont{
         /*
         CustomFont v2.00 (2016-2-24) by tmplinshi
         ---------------------------------------------------------
         Description: Load font from file or resource, without needed install to system.
         ---------------------------------------------------------
         Useage Examples:

            * Load From File
               font1 := New CustomFont("ewatch.ttf")
               Gui, Font, s100, ewatch

            * Load From Resource
               Gui, Add, Text, HWNDhCtrl w400 h200, 12345
               font2 := New CustomFont("res:ewatch.ttf", "ewatch", 80) ; <- Add a res: prefix to the resource name.
               font2.ApplyTo(hCtrl)

            * The fonts will removed automatically when script exits.
              To remove a font manually, just clear the variable (e.g. font1 := "").
         */

         static FR_PRIVATE  := 0x10

         __New(FontFile, FontName="", FontSize=30) {
            if RegExMatch(FontFile, "i)res:\K.*", _FontFile) {
               this.AddFromResource(_FontFile, FontName, FontSize)
            } else {
               this.AddFromFile(FontFile)
            }
         }

         AddFromFile(FontFile) {
            DllCall( "AddFontResourceEx", "Str", FontFile, "UInt", this.FR_PRIVATE, "UInt", 0 )
            this.data := FontFile
         }

         AddFromResource(ResourceName, FontName, FontSize = 30) {
            static FW_NORMAL := 400, DEFAULT_CHARSET := 0x1

            nSize    := this.ResRead(fData, ResourceName)
            fh       := DllCall( "AddFontMemResourceEx", "Ptr", &fData, "UInt", nSize, "UInt", 0, "UIntP", nFonts )
            hFont    := DllCall( "CreateFont", Int,FontSize, Int,0, Int,0, Int,0, UInt,FW_NORMAL, UInt,0
                        , Int,0, Int,0, UInt,DEFAULT_CHARSET, Int,0, Int,0, Int,0, Int,0, Str,FontName )

            this.data := {fh: fh, hFont: hFont}
         }

         ApplyTo(hCtrl) {
            SendMessage, 0x30, this.data.hFont, 1,, ahk_id %hCtrl%
         }

         __Delete() {
            if IsObject(this.data) {
               DllCall( "RemoveFontMemResourceEx", "UInt", this.data.fh    )
               DllCall( "DeleteObject"           , "UInt", this.data.hFont )
            } else {
               DllCall( "RemoveFontResourceEx"   , "Str", this.data, "UInt", this.FR_PRIVATE, "UInt", 0 )
            }
         }

         ; ResRead() By SKAN, from http://www.autohotkey.com/board/topic/57631-crazy-scripting-resource-only-dll-for-dummies-36l-v07/?p=609282
         ResRead( ByRef Var, Key ) {
            VarSetCapacity( Var, 128 ), VarSetCapacity( Var, 0 )
            If ! ( A_IsCompiled ) {
               FileGetSize, nSize, %Key%
               FileRead, Var, *c %Key%
               Return nSize
            }

            If hMod := DllCall( "GetModuleHandle", UInt,0 )
               If hRes := DllCall( "FindResource", UInt,hMod, Str,Key, UInt,10 )
                  If hData := DllCall( "LoadResource", UInt,hMod, UInt,hRes )
                     If pData := DllCall( "LockResource", UInt,hData )
                        Return VarSetCapacity( Var, nSize := DllCall( "SizeofResource", UInt,hMod, UInt,hRes ) )
                           ,  DllCall( "RtlMoveMemory", Str,Var, UInt,pData, UInt,nSize )
            Return 0
         }
      }

      class Image{

         ScreenWidth := A_ScreenWidth, ScreenHeight := A_ScreenHeight

         __New(name := "") {
            this.name := name := (name == "") ? Vis2.Graphics.Name() "_Graphics_Image" : name "_Graphics_Image"

            Vis2.Graphics.Startup()
            Gui, %name%: New, +LastFound +AlwaysOnTop -Caption -DPIScale +E0x80000 +ToolWindow +hwndSecretName, % this.name
            this.hwnd := SecretName
            DllCall("ShowWindow", "ptr",this.hwnd, "int",8)
            this.hbm := CreateDIBSection(A_ScreenWidth, A_ScreenHeight)
            this.hdc := CreateCompatibleDC()
            this.obm := SelectObject(this.hdc, this.hbm)
            this.G := Gdip_GraphicsFromHDC(this.hdc)
            Gdip_SetInterpolationMode(this.G, 7)
         }

         __Delete(){
            Vis2.Graphics.Shutdown()
         }

         Border() {
            Gui, % this.name ": +Border"
            UpdateLayeredWindow(this.hwnd, this.hdc, (A_ScreenWidth-this.w)/2, (A_ScreenHeight-this.h)/2, this.w, this.h)
            return this
         }

         Destroy() {
            SelectObject(this.hdc, this.obm)
            DeleteObject(this.hbm)
            DeleteDC(this.hdc)
            Gdip_DeleteGraphics(this.G)
            Gui, % this.name ":Destroy"
            return this
         }

         Hide() {
            Gui, % this.name ":Show", Hide
            return this
         }

         Show() {
            Gui, % this.name ":Show", NoActivate
            return this
         }

         ToggleVisible() {
            if DllCall("IsWindowVisible", "UInt", this.hwnd)
               Gui, % this.name ":Show", Hide
            else
               Gui, % this.name ":Show", NoActivate
            return this
         }

         isVisible() {
            return DllCall("IsWindowVisible", "UInt", this.hwnd)
         }

         Render(file, scale := 1, time := 0) {
            if (this.hwnd){
               this.scale := scale
               if (time) {
                  self_destruct := ObjBindMethod(this, "Destroy")
                  SetTimer, % self_destruct, % -1 * time
               }
               Critical On
               if FileExist(file)
                  f := pBitmap := Gdip_CreateBitmapFromFile(file)
               else pBitmap := file
               Width := Gdip_GetImageWidth(pBitmap), Height := Gdip_GetImageHeight(pBitmap)
               this.DetectScreenResolutionChange(Width, Height)
               Gdip_DrawImage(this.G, pBitmap, 0, 0, Floor(Width*scale), Floor(Height*scale), 0, 0, Width, Height)
               UpdateLayeredWindow(this.hwnd, this.hdc, 0, 0, Floor(Width*scale), Floor(Height*scale))
               this.w := Floor(Width*scale)
               this.h := Floor(Height*scale)
               if f
                  Gdip_DisposeImage(pBitmap)
               Critical Off
               return this
            }
            else {
               parent := ((___ := RegExReplace(A_ThisFunc, "^(.*)\..*\..*$", "$1")) != A_ThisFunc) ? ___ : ""
               Loop, Parse, parent, .
                  parent := (A_Index=1) ? %A_LoopField% : parent[A_LoopField]
               _image := (parent) ? new parent.image() : new image()
               return _image.Render(file, scale, time)
            }
         }

         DetectScreenResolutionChange(w:="", h:=""){
            w := (w) ? w : A_ScreenWidth
            h := (h) ? h : A_ScreenHeight
            if (this.ScreenWidth != w || this.ScreenHeight != h) {
               this.ScreenWidth := w, this.ScreenHeight := h
               SelectObject(this.hdc, this.obm)
               DeleteObject(this.hbm)
               DeleteDC(this.hdc)
               Gdip_DeleteGraphics(this.G)
               this.hbm := CreateDIBSection(this.ScreenWidth, this.ScreenHeight)
               this.hdc := CreateCompatibleDC()
               this.obm := SelectObject(this.hdc, this.hbm)
               this.G := Gdip_GraphicsFromHDC(this.hdc)
               Gdip_SetInterpolationMode(this.G, 7)
            }
         }
      }

      class Subtitle{

         layers := {}, ScreenWidth := A_ScreenWidth, ScreenHeight := A_ScreenHeight

         __New(name := ""){
            parent := ((___ := RegExReplace(A_ThisFunc, "^(.*)\..*\..*$", "$1")) != A_ThisFunc) ? ___ : ""
            Loop, Parse, parent, .
               this.parent := (A_Index=1) ? %A_LoopField% : this.parent[A_LoopField]

            this.parent.Startup()
            Gui, New, +LastFound +AlwaysOnTop -Caption -DPIScale +E0x80000 +ToolWindow +hwndSecretName
            this.hwnd := SecretName
            this.name := (name != "") ? name "_Subtitle" : "Subtitle_" this.hwnd
            DllCall("ShowWindow", "ptr",this.hwnd, "int",8)
            DllCall("SetWindowText", "ptr",this.hwnd, "str",this.name)
            this.hbm := CreateDIBSection(this.ScreenWidth, this.ScreenHeight)
            this.hdc := CreateCompatibleDC()
            this.obm := SelectObject(this.hdc, this.hbm)
            this.G := Gdip_GraphicsFromHDC(this.hdc)
            this.colorMap := this.colorMap()
         }

         __Delete(){
            this.parent.Shutdown()
         }

         FreeMemory(){
            SelectObject(this.hdc, this.obm)
            DeleteObject(this.hbm)
            DeleteDC(this.hdc)
            Gdip_DeleteGraphics(this.G)
            return this
         }

         Destroy(){
            this.FreeMemory()
            DllCall("DestroyWindow", "ptr",this.hwnd)
            return this
         }

         Hide(){
            DllCall("ShowWindow", "ptr",this.hwnd, "int",0)
            return this
         }

         Show(){
            DllCall("ShowWindow", "ptr",this.hwnd, "int",8)
            return this
         }

         ToggleVisible(){
            this.isVisible() ? this.Hide() : this.Show()
            return this
         }

         isVisible(){
            return DllCall("IsWindowVisible", "ptr",this.hwnd)
         }

         ClickThrough(){
            DetectHiddenWindows On
            WinSet, ExStyle, +0x20, % "ahk_id" this.hwnd
            DetectHiddenWindows Off
            return this
         }

         DetectScreenResolutionChange(w:="", h:=""){
            w := (w) ? w : A_ScreenWidth
            h := (h) ? h : A_ScreenHeight
            if (this.ScreenWidth != w || this.ScreenHeight != h) {
               this.ScreenWidth := w, this.ScreenHeight := h
               SelectObject(this.hdc, this.obm)
               DeleteObject(this.hbm)
               DeleteDC(this.hdc)
               Gdip_DeleteGraphics(this.G)
               this.hbm := CreateDIBSection(this.ScreenWidth, this.ScreenHeight)
               this.hdc := CreateCompatibleDC()
               this.obm := SelectObject(this.hdc, this.hbm)
               this.G := Gdip_GraphicsFromHDC(this.hdc)
            }
         }

         Draw(text := "", style1 := "", style2 := "", pGraphics := "") {
            if (pGraphics == "") {
               pGraphics := this.G
               if (this.rendered == true) {
                  this.rendered := false
                  this.layers := {}
                  this.x := this.y := this.xx := this.yy := ""
                  Gdip_GraphicsClear(this.G)
               }
               this.layers.push([text, style1, style2])
            }

            ; Remember styles so that they can be loaded next time.
            style1 := (style1) ? this.style1 := style1 : this.style1
            style2 := (style2) ? this.style2 := style2 : this.style2

            static q1 := "i)^.*?(?<!-|:|:\s)\b(?![^\(]*\))"
            static q2 := "(:\s?)?\(?(?<value>(?<=\()[\s\-\da-z\.#%]+(?=\))|[\-\da-z\.#%]+).*$"

            time := (style1.t) ? style1.t : (style1.time) ? style1.time
                  : (!IsObject(style1) && (___ := RegExReplace(style1, q1 "(t(ime)?)" q2, "${value}")) != style1) ? ___
                  : (style2.t) ? style2.t : (style2.time) ? style2.time
                  : (!IsObject(style2) && (___ := RegExReplace(style2, q1 "(t(ime)?)" q2, "${value}")) != style2) ? ___
                  : 0

            if (time) {
               self_destruct := ObjBindMethod(this, "Destroy")
               SetTimer, % self_destruct, % -1 * time
            }

            static alpha := "^[A-Za-z]+$"
            static decimal := "^(\-?\d+(\.\d*)?)$"
            static integer := "^\d+$"
            static percentage := "^(\-?\d+(?:\.\d*)?)%$"
            static positive := "^\d+(\.\d*)?$"

            if IsObject(style1){
               _a  := (style1.a != "")  ? style1.a  : style1.anchor
               _x  := (style1.x != "")  ? style1.x  : style1.left
               _y  := (style1.y != "")  ? style1.y  : style1.top
               _w  := (style1.w != "")  ? style1.w  : style1.width
               _h  := (style1.h != "")  ? style1.h  : style1.height
               _r  := (style1.r != "")  ? style1.r  : style1.radius
               _c  := (style1.c != "")  ? style1.c  : style1.color
               _m  := (style1.m != "")  ? style1.m  : style1.margin
               _p  := (style1.p != "")  ? style1.p  : style1.padding
               _q  := (style1.q != "")  ? style1.q  : (style1.quality) ? style1.quality : style1.SmoothingMode
            } else {
               _a  := ((___ := RegExReplace(style1, q1    "(a(nchor)?)"            q2, "${value}")) != style1) ? ___ : ""
               _x  := ((___ := RegExReplace(style1, q1    "(x|left)"               q2, "${value}")) != style1) ? ___ : ""
               _y  := ((___ := RegExReplace(style1, q1    "(y|top)"                q2, "${value}")) != style1) ? ___ : ""
               _w  := ((___ := RegExReplace(style1, q1    "(w(idth)?)"             q2, "${value}")) != style1) ? ___ : ""
               _h  := ((___ := RegExReplace(style1, q1    "(h(eight)?)"            q2, "${value}")) != style1) ? ___ : ""
               _r  := ((___ := RegExReplace(style1, q1    "(r(adius)?)"            q2, "${value}")) != style1) ? ___ : ""
               _c  := ((___ := RegExReplace(style1, q1    "(c(olor)?)"             q2, "${value}")) != style1) ? ___ : ""
               _m  := ((___ := RegExReplace(style1, q1    "(m(argin)?)"            q2, "${value}")) != style1) ? ___ : ""
               _p  := ((___ := RegExReplace(style1, q1    "(p(adding)?)"           q2, "${value}")) != style1) ? ___ : ""
               _q  := ((___ := RegExReplace(style1, q1    "(q(uality)?)"           q2, "${value}")) != style1) ? ___ : ""
            }

            if IsObject(style2){
               a  := (style2.a != "")  ? style2.a  : style2.anchor
               x  := (style2.x != "")  ? style2.x  : style2.left
               y  := (style2.y != "")  ? style2.y  : style2.top
               w  := (style2.w != "")  ? style2.w  : style2.width
               h  := (style2.h != "")  ? style2.h  : style2.height
               m  := (style2.m != "")  ? style2.m  : style2.margin
               f  := (style2.f != "")  ? style2.f  : style2.font
               s  := (style2.s != "")  ? style2.s  : style2.size
               c  := (style2.c != "")  ? style2.c  : style2.color
               b  := (style2.b != "")  ? style2.b  : style2.bold
               i  := (style2.i != "")  ? style2.i  : style2.italic
               u  := (style2.u != "")  ? style2.u  : style2.underline
               j  := (style2.j != "")  ? style2.j  : style2.justify
               n  := (style2.n != "")  ? style2.n  : style2.noWrap
               z  := (style2.z != "")  ? style2.z  : style2.condensed
               d  := (style2.d != "")  ? style2.d  : style2.dropShadow
               o  := (style2.o != "")  ? style2.o  : style2.outline
               q  := (style2.q != "")  ? style2.q  : (style2.quality) ? style2.quality : style2.TextRenderingHint
            } else {
               a  := ((___ := RegExReplace(style2, q1    "(a(nchor)?)"            q2, "${value}")) != style2) ? ___ : ""
               x  := ((___ := RegExReplace(style2, q1    "(x|left)"               q2, "${value}")) != style2) ? ___ : ""
               y  := ((___ := RegExReplace(style2, q1    "(y|top)"                q2, "${value}")) != style2) ? ___ : ""
               w  := ((___ := RegExReplace(style2, q1    "(w(idth)?)"             q2, "${value}")) != style2) ? ___ : ""
               h  := ((___ := RegExReplace(style2, q1    "(h(eight)?)"            q2, "${value}")) != style2) ? ___ : ""
               m  := ((___ := RegExReplace(style2, q1    "(m(argin)?)"            q2, "${value}")) != style2) ? ___ : ""
               f  := ((___ := RegExReplace(style2, q1    "(f(ont)?)"              q2, "${value}")) != style2) ? ___ : ""
               s  := ((___ := RegExReplace(style2, q1    "(s(ize)?)"              q2, "${value}")) != style2) ? ___ : ""
               c  := ((___ := RegExReplace(style2, q1    "(c(olor)?)"             q2, "${value}")) != style2) ? ___ : ""
               b  := ((___ := RegExReplace(style2, q1    "(b(old)?)"              q2, "${value}")) != style2) ? ___ : ""
               i  := ((___ := RegExReplace(style2, q1    "(i(talic)?)"            q2, "${value}")) != style2) ? ___ : ""
               u  := ((___ := RegExReplace(style2, q1    "(u(nderline)?)"         q2, "${value}")) != style2) ? ___ : ""
               j  := ((___ := RegExReplace(style2, q1    "(j(ustify)?)"           q2, "${value}")) != style2) ? ___ : ""
               n  := ((___ := RegExReplace(style2, q1    "(n(oWrap)?)"            q2, "${value}")) != style2) ? ___ : ""
               z  := ((___ := RegExReplace(style2, q1    "(z|condensed?)"         q2, "${value}")) != style2) ? ___ : ""
               d  := ((___ := RegExReplace(style2, q1    "(d(ropShadow)?)"        q2, "${value}")) != style2) ? ___ : ""
               o  := ((___ := RegExReplace(style2, q1    "(o(utline)?)"           q2, "${value}")) != style2) ? ___ : ""
               q  := ((___ := RegExReplace(style2, q1    "(q(uality)?)"           q2, "${value}")) != style2) ? ___ : ""
            }

            ; Step 1 - Simulate string width and height, setting only the variables we need to determine it.
            style += (b) ? 1 : 0      ; bold
            style += (i) ? 2 : 0      ; italic
            style += (u) ? 4 : 0      ; underline
            style += (strike) ? 8 : 0 ; strikeout, not implemented.
            s  := (s ~= percentage) ? A_ScreenHeight * SubStr(s, 1, -1)  / 100 :  s
            s  := (s ~= positive) ? s : 36
            q  := (q >= 0 && q <= 5) ? q : 4
            n  := (n) ? 0x4000 | 0x1000 : 0x4000
            j  := (j ~= "i)cent(er|re)") ? 1 : (j ~= "i)(far|right)") ? 2 : 0
            _q := (_q >= 0 && _q <= 4) ? _q : 4

            Gdip_SetSmoothingMode(pGraphics, _q)
            Gdip_SetTextRenderingHint(pGraphics, q) ; 4 = Anti-Alias, 5 = Cleartype
            hFamily := (___ := Gdip_FontFamilyCreate(f)) ? ___ : Gdip_FontFamilyCreate("Arial")
            hFont := Gdip_FontCreate(hFamily, s, style)
            hFormat := Gdip_StringFormatCreate(n)
            Gdip_SetStringFormatAlign(hFormat, j)

            CreateRectF(RC, 0, 0, 0, 0)
            ReturnRC := Gdip_MeasureString(pGraphics, Text, hFont, hFormat, RC)
            ReturnRC := StrSplit(ReturnRC, "|")

            ; Step 2 - Define margins and padding.
            _m := this.margin(_m)
            _p := this.margin(_p)
             m := this.margin( m)
             p := this.margin( p)

            ; Bonus - Condense Text using a Condensed Font if simulated text width exceeds screen width.
            if (___ := Gdip_FontFamilyCreate(z)) {
               ExtraMargin := (_m.2 + _m.4 + _p.2 + _p.4)
               if (ReturnRC[3] + ExtraMargin > A_ScreenWidth){
                  hFamily := Gdip_FontFamilyCreate(z)
                  hFont := Gdip_FontCreate(hFamily, s, style)
                  ReturnRC := Gdip_MeasureString(pGraphics, Text, hFont, hFormat, RC)
                  ReturnRC := StrSplit(ReturnRC, "|")
                  _w  := ReturnRC[3]
               }
            }

            ; Step 3 - Define _width and _height. Do not modify with margin and padding.
            _w  := (_w  ~= percentage) ? A_ScreenWidth  * SubStr(_w, 1, -1)  / 100 : _w
            _h  := (_h  ~= percentage) ? A_ScreenHeight * SubStr(_h, 1, -1)  / 100 : _h
            _w  := (_w  ~= positive) ? _w  : ReturnRC[3]
            _h  := (_h  ~= positive) ? _h  : ReturnRC[4]

            ; Step 4 - Define _anchor with a default value of 1.
            _a  := (_a = "top") ? 2 : (_a = "left") ? 4 : (_a = "right") ? 6 : (_a = "bottom") ? 8
                  : (_a ~= "i)top" && _a ~= "i)left") ? 1 : (_a ~= "i)top" && _a ~= "i)cent(er|re)") ? 2
                  : (_a ~= "i)top" && _a ~= "i)bottom") ? 3 : (_a ~= "i)cent(er|re)" && _a ~= "i)left") ? 4
                  : (_a ~= "i)cent(er|re)") ? 5 : (_a ~= "i)cent(er|re)" && _a ~= "i)bottom") ? 6
                  : (_a ~= "i)bottom" && _a ~= "i)left") ? 7 : (_a ~= "i)bottom" && _a ~= "i)cent(er|re)") ? 8
                  : (_a ~= "i)bottom" && _a ~= "i)right") ? 9 : (_a ~= "^[1-9]$") ? _a : 1 ; default

            ; Step 5 - Modify _anchor with _x and _y.
            _a  := (_x  = "left") ? 1+(((_a-1)//3)*3) : (_x ~= "i)cent(er|re)") ? 2+(((_a-1)//3)*3) : (_x = "right") ? 3+(((_a-1)//3)*3) : _a
            _a  := (_y  = "top") ? 1+(mod(_a-1,3)) : (_y ~= "i)cent(er|re)") ? 4+(mod(_a-1,3)) : (_y = "bottom") ? 7+(mod(_a-1,3)) : _a

            ; Step 6 - Define _x and _y with respect to _anchor.
            _x  := (_x  = "left") ? 0 : (_x ~= "i)cent(er|re)") ? 0.5*A_ScreenWidth : (_x = "right") ? A_ScreenWidth : _x
            _y  := (_y  = "top") ? 0 : (_y ~= "i)cent(er|re)") ? 0.5*A_ScreenHeight : (_y = "bottom") ? A_ScreenHeight : _y
            _x  := (_x  ~= percentage) ? A_ScreenWidth  * SubStr(_x, 1, -1)  / 100 : _x
            _y  := (_y  ~= percentage) ? A_ScreenHeight * SubStr(_y, 1, -1)  / 100 : _y
            _x  := (_x  ~= decimal) ? _x  : 0
            _y  := (_y  ~= decimal) ? _y  : 0
            _x  -= (mod(_a-1,3) == 0) ? 0 : (mod(_a-1,3) == 1) ? _w/2 : (mod(_a-1,3) == 2) ? _w : 0
            _y  -= (((_a-1)//3) == 0) ? 0 : (((_a-1)//3) == 1) ? _h/2 : (((_a-1)//3) == 2) ? _h : 0
            ; Fractional y values might cause gdi+ slowdown.


            ; Round 1 - Define width and height.
            w  := ( w  ~= percentage) ? _w * RegExReplace( w, percentage, "$1")  / 100 : w
            h  := ( h  ~= percentage) ? _h * RegExReplace( h, percentage, "$1")  / 100 : h
            w  := ( w  ~= positive) ?  w  : (_w) ? _w : ReturnRC[3] ;if _w = 0
            h  := ( h  ~= positive) ?  h  : (_h) ? _h : ReturnRC[4]

            ; Round 2 - Define anchor.
            a  := (a = "top") ? 2 : (a = "left") ? 4 : (a = "right") ? 6 : (a = "bottom") ? 8
                  : (a ~= "i)top" && a ~= "i)left") ? 1 : (a ~= "i)top" && a ~= "i)cent(er|re)") ? 2
                  : (a ~= "i)top" && a ~= "i)bottom") ? 3 : (a ~= "i)cent(er|re)" && a ~= "i)left") ? 4
                  : (a ~= "i)cent(er|re)") ? 5 : (_a ~= "i)cent(er|re)" && a ~= "i)bottom") ? 6
                  : (a ~= "i)bottom" && a ~= "i)left") ? 7 : (a ~= "i)bottom" && a ~= "i)cent(er|re)") ? 8
                  : (a ~= "i)bottom" && a ~= "i)right") ? 9 : (a ~= "^[1-9]$") ? a : 1 ; default

            a  := ( x  = "left") ? 1+((( a-1)//3)*3) : ( x ~= "i)cent(er|re)") ? 2+((( a-1)//3)*3) : ( x = "right") ? 3+((( a-1)//3)*3) :  a
            a  := ( y  = "top") ? 1+(mod( a-1,3)) : ( y ~= "i)cent(er|re)") ? 4+(mod( a-1,3)) : ( y = "bottom") ? 7+(mod( a-1,3)) :  a

            ; Round 3 - Define x and y with respect to anchor.
            x  := ( x  = "left") ? _x : (x ~= "i)cent(er|re)") ? _x + 0.5*_w : (x = "right") ? _x + _w : x
            y  := ( y  = "top") ? _y : (y ~= "i)cent(er|re)") ? _y + 0.5*_h : (y = "bottom") ? _y + _h : y
            x  := ( x  ~= percentage) ? _x + (_w * RegExReplace( x, percentage, "$1")  / 100) : x
            y  := ( y  ~= percentage) ? _y + (_h * RegExReplace( y, percentage, "$1")  / 100) : y
            x  := ( x  ~= decimal) ? x  : _x
            y  := ( y  ~= decimal) ? y  : _y
            x  -= (mod(a-1,3) == 0) ? 0 : (mod(a-1,3) == 1) ? ReturnRC[3]/2 : (mod(a-1,3) == 2) ? ReturnRC[3] : 0
            y  -= (((a-1)//3) == 0) ? 0 : (((a-1)//3) == 1) ? ReturnRC[4]/2 : (((a-1)//3) == 2) ? ReturnRC[4] : 0

            ; Round 4 - Modify _x, _y, _w, _h with margin and padding.
            if (_w && _h) {
               _w  += (_m.2 + _m.4 + _p.2 + _p.4) + (m.2 + m.4 + p.2 + p.4)
               _h  += (_m.1 + _m.3 + _p.1 + _p.3) + (m.1 + m.3 + p.1 + p.3)
               _x  -= (_m.1 + _p.1)
               _y  -= (_m.4 + _p.4)
            }

            ; Round 5 - Modify x, y with margin and padding.
            x  += (m.1 + p.1)
            y  += (m.4 + p.4)

            ; Round 6 - Define radius of rounded corners.
            _smaller := (_w > _h) ? _h : _w
            _r  := (_r  ~= percentage) ? _smaller * RegExReplace(_r, percentage, "$1")  / 100 : _r
            _r  := (_r  <= _smaller / 2 && _r ~= positive) ? _r : 0

            ; Round 7 - Define color.
            _c := this.color(_c, 0xDD424242)
             c := this.color( c, 0xFFFFFFFF)

            ; Round 8 - Define outline and dropShadow.
            o := this.outline(o)
            d := this.dropShadow(d)

            ; Round 9 - Define Text
            if (!A_IsUnicode){
               nSize := DllCall("MultiByteToWideChar", "uint",0, "uint",0, "ptr",&text, "int",-1, "ptr",0, "int",0)
               VarSetCapacity(wtext, nSize*2)
               DllCall("MultiByteToWideChar", "uint",0, "uint",0, "ptr",&text, "int",-1, "ptr",&wtext, "int",nSize)
            }


            ; Draw 1 - Background
            if (_w && _h && _c && (_c & 0xFF000000)) {
               pBrushBackground := Gdip_BrushCreateSolid(_c)
               Gdip_FillRoundedRectangle(pGraphics, pBrushBackground, _x, _y, _w, _h, _r)
               Gdip_DeleteBrush(pBrushBackground)
            }

            ; Draw 2 - DropShadow
            if (!d.void) {
               delta := 2*d.3 + 2*o.1
               offset := d.3 + o.1

               if (d.3) {
                  pBitmap := Gdip_CreateBitmap(w + delta, h + delta)
                  pGraphicsDropShadow := Gdip_GraphicsFromImage(pBitmap)
                  Gdip_SetSmoothingMode(pGraphicsDropShadow, _q)
                  Gdip_SetTextRenderingHint(pGraphicsDropShadow, q)
                  CreateRectF(RC, offset, offset, w + delta, h + delta)
               } else {
                  CreateRectF(RC, x + d.1, y + d.2, w, h)
                  pGraphicsDropShadow := pGraphics
               }

               if (!o.void)
               {
                  DllCall("gdiplus\GdipCreatePath", "int",1, "uptr*",pPath)
                  DllCall("gdiplus\GdipAddPathString", "ptr",pPath, "ptr", A_IsUnicode ? &text : &wtext, "int",-1
                                                     , "ptr",hFamily, "int",style, "float",s, "ptr",&RC, "ptr",hFormat)
                  pPen := Gdip_CreatePen(d.4, o.1)
                  DllCall("gdiplus\GdipSetPenLineJoin", "ptr",pPen, "uInt",2)
                  DllCall("gdiplus\GdipDrawPath", "ptr",pGraphicsDropShadow, "ptr",pPen, "ptr",pPath)
                  Gdip_DeletePen(pPen)
                  pBrush := Gdip_BrushCreateSolid(d.4)
                  Gdip_SetCompositingMode(pGraphicsDropShadow, 1) ; Turn off alpha blending
                  Gdip_SetSmoothingMode(pGraphicsDropShadow, 3)   ; Turn off anti-aliasing
                  Gdip_FillPath(pGraphicsDropShadow, pBrush, pPath)
                  Gdip_DeleteBrush(pBrush)
                  Gdip_DeletePath(pPath)
                  Gdip_SetCompositingMode(pGraphicsDropShadow, 0)
                  Gdip_SetSmoothingMode(pGraphicsDropShadow, _q)
               }
               else
               {
                  pBrush := Gdip_BrushCreateSolid(d.4)
                  Gdip_DrawString(pGraphicsDropShadow, Text, hFont, hFormat, pBrush, RC)
                  Gdip_DeleteBrush(pBrush)
               }

               if (d.3) {
                  Gdip_DeleteGraphics(pGraphicsDropShadow)
                  pBlur := Gdip_BlurBitmap(pBitmap, d.3)
                  Gdip_DisposeImage(pBitmap)
                  Gdip_DrawImage(pGraphics, pBlur, x + d.1 - offset, y + d.2 - offset, w + delta, h + delta)
                  Gdip_DisposeImage(pBlur)
               }
            }

            ; Draw 3 - Text Outline
            if (!o.void) {
               ; Convert our text to a path.
               CreateRectF(RC, x, y, w, h)
               DllCall("gdiplus\GdipCreatePath", "int",1, "uptr*",pPath)
               DllCall("gdiplus\GdipAddPathString", "ptr",pPath, "ptr", A_IsUnicode ? &text : &wtext, "int",-1
                                                  , "ptr",hFamily, "int",style, "float",s, "ptr",&RC, "ptr",hFormat)

               ; Create a pen.
               pPen := Gdip_CreatePen(o.2, o.1)
               DllCall("gdiplus\GdipSetPenLineJoin", "ptr",pPen, "uint",2)

               ; Create a glow effect around the edges.
               if (o.3) {
                  DllCall("gdiplus\GdipClonePath", "ptr",pPath, "uptr*",pPathGlow)
                  DllCall("gdiplus\GdipWidenPath", "ptr",pPathGlow, "ptr",pPen, "ptr",0, "float",1)

                  ; Set color to glowColor or use the previous color.
                  color := (o.4) ? o.4 : o.2

                  loop % o.3
                  {
                     ARGB := Format("0x{:02X}",((color & 0xFF000000) >> 24)/o.3) . Format("{:06X}",(color & 0x00FFFFFF))
                     pPenGlow := Gdip_CreatePen(ARGB, A_Index)
                     DllCall("gdiplus\GdipSetPenLineJoin", "ptr",pPenGlow, "uInt",2)
                     DllCall("gdiplus\GdipDrawPath", "ptr",pGraphics, "ptr",pPenGlow, "ptr",pPathGlow)
                     Gdip_DeletePen(pPenGlow)
                  }
                  Gdip_DeletePath(pPathGlow)
               }

               ; Draw outline text.
               if (o.1)
                  DllCall("gdiplus\GdipDrawPath", "ptr",pGraphics, "ptr",pPen, "ptr",pPath)

               ; Fill outline text.
               if (c && (c & 0xFF000000)) {
                  pBrush := Gdip_BrushCreateSolid(c)
                  Gdip_FillPath(pGraphics, pBrush, pPath)
                  Gdip_DeleteBrush(pBrush)
               }
               Gdip_DeletePen(pPen)
               Gdip_DeletePath(pPath)
            }

            ; Draw Text
            if (text != "" && d.void && o.void) {
               CreateRectF(RC, x, y, w, h)
               pBrushText := Gdip_BrushCreateSolid(c)
               Gdip_DrawString(pGraphics, text, hFont, hFormat, pBrushText, RC)
               Gdip_DeleteBrush(pBrushText)
            }

            ; Complete
            Gdip_DeleteStringFormat(hFormat)
            Gdip_DeleteFont(hFont)
            Gdip_DeleteFontFamily(hFamily)

            ; Correct Offsets
            _w := (_w == 0) ? (ReturnRC[3] + d.1 + 2*d.3 + 2*o.1 + 2*o.3) : _w
            _h := (_h == 0) ? (ReturnRC[4] + d.2 + 2*d.3 + 2*o.1 + 2*o.3) : _h

            this.x  := (this.x  = "" || _x < this.x) ? _x : this.x
            this.y  := (this.y  = "" || _y < this.y) ? _y : this.y
            this.xx := (this.xx = "" || _x + _w > this.xx) ? _x + _w : this.xx
            this.yy := (this.yy = "" || _y + _h > this.yy) ? _y + _h : this.yy
            return
         }

         Render(text := "", style1 := "", style2 := "", update := 1){
            if (this.hWnd){
               Critical On
               this.DetectScreenResolutionChange()
               this.Draw(text, style1, style2)
               if (update)
                  UpdateLayeredWindow(this.hwnd, this.hdc, 0, 0, A_ScreenWidth, A_ScreenHeight)
               this.rendered := true
               Critical Off
               return this
            }
            else {
               parent := ((___ := RegExReplace(A_ThisFunc, "^(.*)\..*\..*$", "$1")) != A_ThisFunc) ? ___ : ""
               Loop, Parse, parent, .
                  parent := (A_Index=1) ? %A_LoopField% : parent[A_LoopField]
               _subtitle := (parent) ? new parent.Subtitle() : new Subtitle()
               return _subtitle.Render(text, style1, style2, update)
            }
         }

         Bitmap(x:=0, y:=0, w:=0, h:=0){
            pBitmap := Gdip_CreateBitmap(A_ScreenWidth, A_ScreenHeight)
            pGraphics := Gdip_GraphicsFromImage(pBitmap)
            loop % this.layers.MaxIndex()
               this.Draw(this.layers[A_Index].1, this.layers[A_Index].2, this.layers[A_Index].3, pGraphics)
            Gdip_DeleteGraphics(pGraphics)

            if (x || y || w || h) {
               w := (w = 0) ? A_ScreenWidth, h := (h = 0) ? A_ScreenHeight
               pBitmap2 := Gdip_CloneBitmapArea(pBitmap, x, y, w, h)
               Gdip_DisposeImage(pBitmap)
               pBitmap := pBitmap2
            }
            return pBitmap ; Please dispose of this image responsibly.
         }

         Save(filename := "", quality := 92, fullscreen := 0){
            filename := (filename ~= "i)\.(bmp|dib|rle|jpg|jpeg|jpe|jfif|gif|tif|tiff|png)$") ? filename
                      : (filename != "") ? filename ".png" : this.name ".png"
            pBitmap := (fullscreen) ? this.Bitmap() : this.Bitmap(this.x, this.y, this.xx - this.x, this.yy - this.y)
            Gdip_SaveBitmapToFile(pBitmap, filename, quality)
            Gdip_DisposeImage(pBitmap)
         }

         SaveFullScreen(filename := "", quality := ""){
            return this.Save(filename, quality, 1)
         }

         hBitmap(alpha := 0xFFFFFFFF){
            ; hBitmap converts alpha channel to specified alpha color.
            ; Add 1 pixel because Anti-Alias (SmoothingMode = 4)
            ; Should it be crop 1 pixel instead?
            pBitmap := this.Bitmap(this.x, this.y, this.xx - this.x, this.yy - this.y)
            hBitmap := Gdip_CreateHBITMAPFromBitmap(pBitmap, alpha)
            Gdip_DisposeImage(pBitmap)
            return hBitmap
         }

         RenderToHBitmap(text := "", style1 := "", style2 := ""){
            if (this.hWnd){
               this.Render(text, style1, style2, 0)
               return this.hBitmap()
            }
            else {
               parent := ((___ := RegExReplace(A_ThisFunc, "^(.*)\..*\..*$", "$1")) != A_ThisFunc) ? ___ : ""
               Loop, Parse, parent, .
                  parent := (A_Index=1) ? %A_LoopField% : parent[A_LoopField]
               _subtitle := (parent) ? new parent.Subtitle() : new Subtitle()
               _subtitle.Render(text, style1, style2, 0)
               return _subtitle.hBitmap() ; Does not return a subtitle object.
            }
         }

         hIcon(){
            pBitmap := this.Bitmap(this.x, this.y, this.xx - this.x, this.yy - this.y)
            hIcon := Gdip_CreateHICONFromBitmap(pBitmap)
            Gdip_DisposeImage(pBitmap)
            return hIcon
         }

         color(c, default := 0xDD424242){
            static colorRGB  := "^0x([0-9A-Fa-f]{6})$"
            static colorARGB := "^0x([0-9A-Fa-f]{8})$"
            static hex6      :=   "^([0-9A-Fa-f]{6})$"
            static hex8      :=   "^([0-9A-Fa-f]{8})$"

            if ObjGetCapacity([c], 1){
               c  := (c ~= "^#") ? SubStr(c, 2) : c
               c  := ((___ := this.colorMap[c]) != "") ? ___ : c
               c  := (c ~= colorRGB) ? "0xFF" RegExReplace(c, colorRGB, "$1") : (c ~= hex8) ? "0x" c : (c ~= hex6) ? "0xFF" c : c
               c  := (c ~= colorARGB) ? c : default
            }
            return (c != "") ? c : default
         }

         margin(m, default := 0){
            static percentage := "^(\-?\d+(?:\.\d*)?)%$"
            static positive := "^\d+(\.\d*)?$"

            if IsObject(m){
               m.1 := (m.y  != "") ? m.y  : m.top
               m.2 := (m.x2 != "") ? m.x2 : m.right
               m.3 := (m.y2 != "") ? m.y2 : m.bottom
               m.4 := (m.x  != "") ? m.x  : m.left
            }
            else if (m) {
               m   := StrSplit(m, " ")
               if (m.length() == 3)
                  m.4 := m.2
               else if (m.length() == 2)
                  m.4 := m.2, m.3 := m.1
               else if (m.length() == 1)
                  m.4 := m.3 := m.2 := m.1, exception := true
               else
                  m.Delete(5, m.MaxIndex())
            }
            else
               return {1:default, 2:default, 3:default, 4:default}

            m.1 := (m.1 ~= percentage) ? A_ScreenHeight * SubStr(m.1, 1, -1)  / 100 : m.1
            m.2 := (m.2 ~= percentage) ? (exception ? A_ScreenHeight : A_ScreenWidth) * SubStr(m.2, 1, -1)  / 100 : m.2
            m.3 := (m.3 ~= percentage) ? A_ScreenHeight * SubStr(m.3, 1, -1)  / 100 : m.3
            m.4 := (m.4 ~= percentage) ? (exception ? A_ScreenHeight : A_ScreenWidth) * SubStr(m.4, 1, -1)  / 100 : m.4

            m.1 := (m.1 ~= positive) ? m.1 : default
            m.2 := (m.2 ~= positive) ? m.2 : default
            m.3 := (m.3 ~= positive) ? m.3 : default
            m.4 := (m.4 ~= positive) ? m.4 : default

            return m
         }

         outline(o){
            static percentage := "^(\-?\d+(?:\.\d*)?)%$"
            static positive := "^\d+(\.\d*)?$"

            if IsObject(o){
               o.1 := (o.w  != "") ? o.w  : o.width
               o.2 := (o.c  != "") ? o.c  : o.color
               o.3 := (o.g  != "") ? o.g  : o.glow
               o.4 := (o.c2 != "") ? o.c2 : o.glowColor
            } else if (o)
               o   := StrSplit(o, " ")
            else
               return {"void":true, 1:0, 2:0, 3:0, 4:0}

            o.1 := (o.1 ~= "px$") ? SubStr(o.1, 1, -2) : o.1
            o.1 := (o.1 ~= percentage) ?  s * RegExReplace(o.1, percentage, "$1")  // 100 : o.1
            o.1 := (o.1 ~= positive) ? o.1 : 1

            o.2 := this.color(o.2, 0xFF000000)

            o.3 := (o.3 ~= "px$") ? SubStr(o.3, 1, -2) : o.3
            o.3 := (o.3 ~= percentage) ?  s * RegExReplace(o.3, percentage, "$1")  // 100 : o.3
            o.3 := (o.3 ~= positive) ? o.3 : 0

            o.4 := this.color(o.4, 0x00000000)
            return o
         }

         dropShadow(d){
            static decimal := "^(\-?\d+(\.\d*)?)$"
            static percentage := "^(\-?\d+(?:\.\d*)?)%$"
            static positive := "^\d+(\.\d*)?$"

            if IsObject(d){
               d.1 := (d.h != "") ? d.h : d.horizontal
               d.2 := (d.v != "") ? d.v : d.vertical
               d.3 := (d.b != "") ? d.b : d.blur
               d.4 := (d.c != "") ? d.c : d.color
               d.5 := (d.s != "") ? d.s : d.strength
            } else if (d)
               d   := StrSplit(d, " ")
            else
               return {"void":true, 1:0, 2:0, 3:0, 4:0, 5:0}

            d.1 := (d.1 ~= "px$") ? SubStr(d.1, 1, -2) : d.1
            d.1 := (d.1 ~= percentage) ? ReturnRC[3] * RegExReplace(d.1, percentage, "$1")  / 100 : d.1
            d.1 := (d.1 ~= decimal) ? d.1 : 0

            d.2 := (d.2 ~= "px$") ? SubStr(d.2, 1, -2) : d.2
            d.2 := (d.2 ~= percentage) ? ReturnRC[4] * RegExReplace(d.2, percentage, "$1")  / 100 : d.2
            d.2 := (d.2 ~= decimal) ? d.2 : 0

            d.3 := (d.3 ~= "px$") ? SubStr(d.3, 1, -2) : d.3
            d.3 := (d.3 ~= percentage) ? s * RegExReplace(d.3, percentage, "$1")  / 100 : d.3
            d.3 := (d.3 ~= positive) ? d.3 : 1

            d.4 := this.color(d.4, 0xFF000000)

            d.5 := (d.5 ~= percentage) ? s * RegExReplace(d.5, percentage, "$1")  / 100 : d.5
            d.5 := (d.5 ~= positive) ? d.5 : 1
            return d
         }

         colorMap(){
            color := [] ; 73 LINES MAX
            color["Clear"] := color["Off"] := color["None"] := color["Transparent"] := "0x00000000"

               color["AliceBlue"]             := "0xFFF0F8FF"
             , color["AntiqueWhite"]          := "0xFFFAEBD7"
             , color["Aqua"]                  := "0xFF00FFFF"
             , color["Aquamarine"]            := "0xFF7FFFD4"
             , color["Azure"]                 := "0xFFF0FFFF"
             , color["Beige"]                 := "0xFFF5F5DC"
             , color["Bisque"]                := "0xFFFFE4C4"
             , color["Black"]                 := "0xFF000000"
             , color["BlanchedAlmond"]        := "0xFFFFEBCD"
             , color["Blue"]                  := "0xFF0000FF"
             , color["BlueViolet"]            := "0xFF8A2BE2"
             , color["Brown"]                 := "0xFFA52A2A"
             , color["BurlyWood"]             := "0xFFDEB887"
             , color["CadetBlue"]             := "0xFF5F9EA0"
             , color["Chartreuse"]            := "0xFF7FFF00"
             , color["Chocolate"]             := "0xFFD2691E"
             , color["Coral"]                 := "0xFFFF7F50"
             , color["CornflowerBlue"]        := "0xFF6495ED"
             , color["Cornsilk"]              := "0xFFFFF8DC"
             , color["Crimson"]               := "0xFFDC143C"
             , color["Cyan"]                  := "0xFF00FFFF"
             , color["DarkBlue"]              := "0xFF00008B"
             , color["DarkCyan"]              := "0xFF008B8B"
             , color["DarkGoldenRod"]         := "0xFFB8860B"
             , color["DarkGray"]              := "0xFFA9A9A9"
             , color["DarkGrey"]              := "0xFFA9A9A9"
             , color["DarkGreen"]             := "0xFF006400"
             , color["DarkKhaki"]             := "0xFFBDB76B"
             , color["DarkMagenta"]           := "0xFF8B008B"
             , color["DarkOliveGreen"]        := "0xFF556B2F"
             , color["DarkOrange"]            := "0xFFFF8C00"
             , color["DarkOrchid"]            := "0xFF9932CC"
             , color["DarkRed"]               := "0xFF8B0000"
             , color["DarkSalmon"]            := "0xFFE9967A"
             , color["DarkSeaGreen"]          := "0xFF8FBC8F"
             , color["DarkSlateBlue"]         := "0xFF483D8B"
             , color["DarkSlateGray"]         := "0xFF2F4F4F"
             , color["DarkSlateGrey"]         := "0xFF2F4F4F"
             , color["DarkTurquoise"]         := "0xFF00CED1"
             , color["DarkViolet"]            := "0xFF9400D3"
             , color["DeepPink"]              := "0xFFFF1493"
             , color["DeepSkyBlue"]           := "0xFF00BFFF"
             , color["DimGray"]               := "0xFF696969"
             , color["DimGrey"]               := "0xFF696969"
             , color["DodgerBlue"]            := "0xFF1E90FF"
             , color["FireBrick"]             := "0xFFB22222"
             , color["FloralWhite"]           := "0xFFFFFAF0"
             , color["ForestGreen"]           := "0xFF228B22"
             , color["Fuchsia"]               := "0xFFFF00FF"
             , color["Gainsboro"]             := "0xFFDCDCDC"
             , color["GhostWhite"]            := "0xFFF8F8FF"
             , color["Gold"]                  := "0xFFFFD700"
             , color["GoldenRod"]             := "0xFFDAA520"
             , color["Gray"]                  := "0xFF808080"
             , color["Grey"]                  := "0xFF808080"
             , color["Green"]                 := "0xFF008000"
             , color["GreenYellow"]           := "0xFFADFF2F"
             , color["HoneyDew"]              := "0xFFF0FFF0"
             , color["HotPink"]               := "0xFFFF69B4"
             , color["IndianRed"]             := "0xFFCD5C5C"
             , color["Indigo"]                := "0xFF4B0082"
             , color["Ivory"]                 := "0xFFFFFFF0"
             , color["Khaki"]                 := "0xFFF0E68C"
             , color["Lavender"]              := "0xFFE6E6FA"
             , color["LavenderBlush"]         := "0xFFFFF0F5"
             , color["LawnGreen"]             := "0xFF7CFC00"
             , color["LemonChiffon"]          := "0xFFFFFACD"
             , color["LightBlue"]             := "0xFFADD8E6"
             , color["LightCoral"]            := "0xFFF08080"
             , color["LightCyan"]             := "0xFFE0FFFF"
             , color["LightGoldenRodYellow"]  := "0xFFFAFAD2"
             , color["LightGray"]             := "0xFFD3D3D3"
             , color["LightGrey"]             := "0xFFD3D3D3"
               color["LightGreen"]            := "0xFF90EE90"
             , color["LightPink"]             := "0xFFFFB6C1"
             , color["LightSalmon"]           := "0xFFFFA07A"
             , color["LightSeaGreen"]         := "0xFF20B2AA"
             , color["LightSkyBlue"]          := "0xFF87CEFA"
             , color["LightSlateGray"]        := "0xFF778899"
             , color["LightSlateGrey"]        := "0xFF778899"
             , color["LightSteelBlue"]        := "0xFFB0C4DE"
             , color["LightYellow"]           := "0xFFFFFFE0"
             , color["Lime"]                  := "0xFF00FF00"
             , color["LimeGreen"]             := "0xFF32CD32"
             , color["Linen"]                 := "0xFFFAF0E6"
             , color["Magenta"]               := "0xFFFF00FF"
             , color["Maroon"]                := "0xFF800000"
             , color["MediumAquaMarine"]      := "0xFF66CDAA"
             , color["MediumBlue"]            := "0xFF0000CD"
             , color["MediumOrchid"]          := "0xFFBA55D3"
             , color["MediumPurple"]          := "0xFF9370DB"
             , color["MediumSeaGreen"]        := "0xFF3CB371"
             , color["MediumSlateBlue"]       := "0xFF7B68EE"
             , color["MediumSpringGreen"]     := "0xFF00FA9A"
             , color["MediumTurquoise"]       := "0xFF48D1CC"
             , color["MediumVioletRed"]       := "0xFFC71585"
             , color["MidnightBlue"]          := "0xFF191970"
             , color["MintCream"]             := "0xFFF5FFFA"
             , color["MistyRose"]             := "0xFFFFE4E1"
             , color["Moccasin"]              := "0xFFFFE4B5"
             , color["NavajoWhite"]           := "0xFFFFDEAD"
             , color["Navy"]                  := "0xFF000080"
             , color["OldLace"]               := "0xFFFDF5E6"
             , color["Olive"]                 := "0xFF808000"
             , color["OliveDrab"]             := "0xFF6B8E23"
             , color["Orange"]                := "0xFFFFA500"
             , color["OrangeRed"]             := "0xFFFF4500"
             , color["Orchid"]                := "0xFFDA70D6"
             , color["PaleGoldenRod"]         := "0xFFEEE8AA"
             , color["PaleGreen"]             := "0xFF98FB98"
             , color["PaleTurquoise"]         := "0xFFAFEEEE"
             , color["PaleVioletRed"]         := "0xFFDB7093"
             , color["PapayaWhip"]            := "0xFFFFEFD5"
             , color["PeachPuff"]             := "0xFFFFDAB9"
             , color["Peru"]                  := "0xFFCD853F"
             , color["Pink"]                  := "0xFFFFC0CB"
             , color["Plum"]                  := "0xFFDDA0DD"
             , color["PowderBlue"]            := "0xFFB0E0E6"
             , color["Purple"]                := "0xFF800080"
             , color["RebeccaPurple"]         := "0xFF663399"
             , color["Red"]                   := "0xFFFF0000"
             , color["RosyBrown"]             := "0xFFBC8F8F"
             , color["RoyalBlue"]             := "0xFF4169E1"
             , color["SaddleBrown"]           := "0xFF8B4513"
             , color["Salmon"]                := "0xFFFA8072"
             , color["SandyBrown"]            := "0xFFF4A460"
             , color["SeaGreen"]              := "0xFF2E8B57"
             , color["SeaShell"]              := "0xFFFFF5EE"
             , color["Sienna"]                := "0xFFA0522D"
             , color["Silver"]                := "0xFFC0C0C0"
             , color["SkyBlue"]               := "0xFF87CEEB"
             , color["SlateBlue"]             := "0xFF6A5ACD"
             , color["SlateGray"]             := "0xFF708090"
             , color["SlateGrey"]             := "0xFF708090"
             , color["Snow"]                  := "0xFFFFFAFA"
             , color["SpringGreen"]           := "0xFF00FF7F"
             , color["SteelBlue"]             := "0xFF4682B4"
             , color["Tan"]                   := "0xFFD2B48C"
             , color["Teal"]                  := "0xFF008080"
             , color["Thistle"]               := "0xFFD8BFD8"
             , color["Tomato"]                := "0xFFFF6347"
             , color["Turquoise"]             := "0xFF40E0D0"
             , color["Violet"]                := "0xFFEE82EE"
             , color["Wheat"]                 := "0xFFF5DEB3"
             , color["White"]                 := "0xFFFFFFFF"
             , color["WhiteSmoke"]            := "0xFFF5F5F5"
               color["Yellow"]                := "0xFFFFFF00"
             , color["YellowGreen"]           := "0xFF9ACD32"
            return color
         }

         x1(){
            return this.x
         }

         y1(){
            return this.y
         }

         x2(){
            return this.xx
         }

         y2(){
            return this.yy
         }

         width(){
            return this.xx - this.x
         }

         height(){
            return this.yy - this.y
         }
      }
   }
}

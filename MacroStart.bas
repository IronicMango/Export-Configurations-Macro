Attribute VB_Name = "MacroStart"
'Written by Ironic Mango Designs
'Released under MIT License

'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
'SOFTWARE.

Option Explicit

Dim swapp As Object

Dim swModel As IModelDoc2

Dim errorMessage As String

Sub main()

Set swapp = Application.SldWorks

Set swModel = swapp.ActiveDoc

errorMessage = Export_Configurations(swModel, True, 1)

If errorMessage <> "Completed without errors." Then
    MsgBox errorMessage, vbOKOnly + vbCritical
End If

End Sub

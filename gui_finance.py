from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.popup import Popup
from kivy.uix.gridlayout import GridLayout
from os.path import sep, expanduser, isdir, dirname

from kivy.garden.filebrowser import FileBrowser

from kivy.app import App
import test
import finance_pull4
#importing python file

#Global Scope

ind_location=0
dash_location=0







#Make an app by deriving from the kivy provided app class
#sma eas MyApp
class FinanceIndicators(App):
	def build(self):
		#Define a grid layout
		self.layout = GridLayout(cols=2, padding=10)

		#add in abutton
		self.button = Button(text="Select Indicators File")

		self.button2 = Button(text="Select Dashboard File")
		self.button3 = Button(text="Run the Application")


		self.mess = Label(text="Welcome to the Finance Application")

		#self.button2 = Button(text="Exit the app")

		self.layout.add_widget(self.mess)

		self.layout.add_widget(self.button)
		self.layout.add_widget(self.button2)
		self.layout.add_widget(self.button3)

		#attach a callback for the button press event

		self.button.bind(on_press=self.onButtonPress)
		self.button2.bind(on_press=self.onButtonPress2)
		self.button3.bind(on_press=self.onButtonPressRunApp)


		
		return self.layout


	






	def onButtonPress(self,button):

		user_path = dirname('Documents')
		layout = FileBrowser(select_string = 'select', favorites=[(user_path,'Documents')])

		layout.bind(
				on_success = self._fbroswer_success,
				on_canceled=self._fbroswer_canceled
			)


		closeButton = Button(text= "Return to Main")

		layout.add_widget(closeButton)

		#intsntiate the modal popup and display
		popup = Popup(title='Application', content = layout)
		#content = (Lablel(text='This is a demo popup')))
		popup.open()

		#attach close button press with popup.dismiss action


		closeButton.bind(on_press=popup.dismiss)



	def onButtonPress2(self,button):

		user_path = dirname('Documents')
		layout = FileBrowser(select_string = 'select', favorites=[(user_path,'Documents')])

		layout.bind(
				on_success = self._fbroswer_success2,
				on_canceled=self._fbroswer_canceled
			)


		closeButton = Button(text= "Return to Main")

		layout.add_widget(closeButton)

		#intsntiate the modal popup and display
		popup = Popup(title='Application', content = layout)
		#content = (Lablel(text='This is a demo popup')))
		popup.open()

		#attach close button press with popup.dismiss action


		closeButton.bind(on_press=popup.dismiss)



	def onButtonPressRunApp(self,button):

		print("printing the location of the files",ind_location, dash_location)








	#this is the function that has the behavior is the file is successfully selected
	#This selects the indicatos file
	def _fbroswer_success(self, instance):
		print("HI, i found your file", instance.selection)
		input_file = str(instance.selection).replace('[', "")
		input_file = str(input_file).replace(']', "")
		input_file = str(input_file).replace("\\", "/")
		input_file = str(input_file).replace("'", "")


		#hard_input = 'C://Users//jprivera//Desktop//provider_productivity_april.xls'

		global ind_location

		ind_location = instance.selection

		print("file from gui OG",input_file)
		

		

		#test.main(instance.selection)



	#This is the function for the output dashboard file
	def _fbroswer_success2(self, instance):
		print("HI, i found your file", instance.selection)
		input_file = str(instance.selection).replace('[', "")
		input_file = str(input_file).replace(']', "")
		input_file = str(input_file).replace("\\", "/")
		input_file = str(input_file).replace("'", "")


		#hard_input = 'C://Users//jprivera//Desktop//provider_productivity_april.xls'

		global dash_location

		dash_location = input_file

		print("file from gui number 2",input_file)
		

		

		#test.main(instance.selection)




	def _fbroswer_canceled(self, instance):
		print ('cancelled, Close self.')
		#self.dismiss()
		#print(dir(self))
		#on_press=self.dismiss











#run the app

if __name__ == '__main__':
	FinanceIndicators().run()
#pragma once
#include <fstream>
#include <string>
#include <wchar.h>
#include <msclr/marshal.h>
#include <msclr/marshal_cppstd.h>
#include <Shlwapi.h>
#include <vector>
#include <string>
#include <msclr/marshal_cppstd.h>
#include <exception>
#pragma comment(lib, "Shlwapi.lib")
#pragma comment(lib, "Shell32.lib")

namespace QuickLauch {

	using namespace System;
	using namespace System::IO;
	using namespace System::ComponentModel;
	using namespace System::Collections;
	using namespace System::Windows::Forms;
	using namespace System::Data;
	using namespace System::Drawing;
	using namespace System::Runtime::InteropServices;
	using namespace System::Diagnostics;
	using namespace msclr::interop;

	/// <summary>
	/// Сводка для MyForm
	/// </summary>
	public ref class MyForm : public System::Windows::Forms::Form
	{
	public:
		MyForm(void)
		{
			InitializeComponent();
		}

	protected:
		/// <summary>
		/// Освободить все используемые ресурсы.
		/// </summary>
		~MyForm()
		{
			if (components)
			{
				delete components;
			}
		}
	private: System::Windows::Forms::ListView^ listView1;
	protected:
	private: System::Windows::Forms::Button^ addBtn;
	private: System::Windows::Forms::Button^ BtnRemove;
	private: System::Windows::Forms::PictureBox^ pictureBox1;
	private: System::Windows::Forms::Label^ label1;


	private:
		/// <summary>
		/// Обязательная переменная конструктора.
		/// </summary>
		System::ComponentModel::Container ^components;

#pragma region Windows Form Designer generated code
		/// <summary>
		/// Требуемый метод для поддержки конструктора — не изменяйте 
		/// содержимое этого метода с помощью редактора кода.
		/// </summary>
		void InitializeComponent(void)
		{
			System::ComponentModel::ComponentResourceManager^ resources = (gcnew System::ComponentModel::ComponentResourceManager(MyForm::typeid));
			this->listView1 = (gcnew System::Windows::Forms::ListView());
			this->addBtn = (gcnew System::Windows::Forms::Button());
			this->BtnRemove = (gcnew System::Windows::Forms::Button());
			this->pictureBox1 = (gcnew System::Windows::Forms::PictureBox());
			this->label1 = (gcnew System::Windows::Forms::Label());
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->pictureBox1))->BeginInit();
			this->SuspendLayout();
			// 
			// listView1
			// 
			this->listView1->BackColor = System::Drawing::Color::White;
			this->listView1->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 7));
			this->listView1->ForeColor = System::Drawing::Color::Black;
			this->listView1->HideSelection = false;
			this->listView1->Location = System::Drawing::Point(276, 12);
			this->listView1->Name = L"listView1";
			this->listView1->Size = System::Drawing::Size(402, 309);
			this->listView1->TabIndex = 0;
			this->listView1->UseCompatibleStateImageBehavior = false;
			this->listView1->DoubleClick += gcnew System::EventHandler(this, &MyForm::listView1_DoubleClick);
			// 
			// addBtn
			// 
			this->addBtn->BackColor = System::Drawing::Color::FromArgb(static_cast<System::Int32>(static_cast<System::Byte>(149)), static_cast<System::Int32>(static_cast<System::Byte>(117)),
				static_cast<System::Int32>(static_cast<System::Byte>(205)));
			this->addBtn->FlatStyle = System::Windows::Forms::FlatStyle::Flat;
			this->addBtn->ForeColor = System::Drawing::Color::White;
			this->addBtn->Location = System::Drawing::Point(565, 336);
			this->addBtn->Name = L"addBtn";
			this->addBtn->Size = System::Drawing::Size(113, 23);
			this->addBtn->TabIndex = 1;
			this->addBtn->Text = L"Добавить";
			this->addBtn->UseVisualStyleBackColor = false;
			this->addBtn->Click += gcnew System::EventHandler(this, &MyForm::addBtn_Click);
			// 
			// BtnRemove
			// 
			this->BtnRemove->BackColor = System::Drawing::Color::FromArgb(static_cast<System::Int32>(static_cast<System::Byte>(149)), static_cast<System::Int32>(static_cast<System::Byte>(117)),
				static_cast<System::Int32>(static_cast<System::Byte>(205)));
			this->BtnRemove->FlatStyle = System::Windows::Forms::FlatStyle::Flat;
			this->BtnRemove->ForeColor = System::Drawing::Color::White;
			this->BtnRemove->Location = System::Drawing::Point(276, 336);
			this->BtnRemove->Name = L"BtnRemove";
			this->BtnRemove->Size = System::Drawing::Size(113, 23);
			this->BtnRemove->TabIndex = 2;
			this->BtnRemove->Text = L"Удалить";
			this->BtnRemove->UseVisualStyleBackColor = false;
			this->BtnRemove->Click += gcnew System::EventHandler(this, &MyForm::BtnRemove_Click);
			// 
			// pictureBox1
			// 
			this->pictureBox1->BackColor = System::Drawing::Color::White;
			this->pictureBox1->ErrorImage = (cli::safe_cast<System::Drawing::Image^>(resources->GetObject(L"pictureBox1.ErrorImage")));
			this->pictureBox1->Image = (cli::safe_cast<System::Drawing::Image^>(resources->GetObject(L"pictureBox1.Image")));
			this->pictureBox1->InitialImage = (cli::safe_cast<System::Drawing::Image^>(resources->GetObject(L"pictureBox1.InitialImage")));
			this->pictureBox1->Location = System::Drawing::Point(12, 12);
			this->pictureBox1->Name = L"pictureBox1";
			this->pictureBox1->Size = System::Drawing::Size(249, 74);
			this->pictureBox1->TabIndex = 3;
			this->pictureBox1->TabStop = false;
			this->pictureBox1->Click += gcnew System::EventHandler(this, &MyForm::pictureBox1_Click);
			// 
			// label1
			// 
			this->label1->Font = (gcnew System::Drawing::Font(L"Nirmala UI", 12));
			this->label1->ForeColor = System::Drawing::Color::White;
			this->label1->Location = System::Drawing::Point(12, 104);
			this->label1->Name = L"label1";
			this->label1->Size = System::Drawing::Size(249, 217);
			this->label1->TabIndex = 4;
			this->label1->Text = resources->GetString(L"label1.Text");
			// 
			// MyForm
			// 
			this->AutoScaleDimensions = System::Drawing::SizeF(6, 13);
			this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
			this->BackColor = System::Drawing::Color::FromArgb(static_cast<System::Int32>(static_cast<System::Byte>(33)), static_cast<System::Int32>(static_cast<System::Byte>(33)),
				static_cast<System::Int32>(static_cast<System::Byte>(33)));
			this->ClientSize = System::Drawing::Size(690, 371);
			this->Controls->Add(this->label1);
			this->Controls->Add(this->pictureBox1);
			this->Controls->Add(this->BtnRemove);
			this->Controls->Add(this->addBtn);
			this->Controls->Add(this->listView1);
			this->Icon = (cli::safe_cast<System::Drawing::Icon^>(resources->GetObject(L"$this.Icon")));
			this->Name = L"MyForm";
			this->Text = L"Quick Launch";
			this->FormClosing += gcnew System::Windows::Forms::FormClosingEventHandler(this, &MyForm::MyForm_FormClosing);
			this->Load += gcnew System::EventHandler(this, &MyForm::MyForm_Load);
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->pictureBox1))->EndInit();
			this->ResumeLayout(false);

		}
#pragma endregion,
	private:
		System::Collections::Generic::Dictionary<String^, String^>^ appPaths = gcnew System::Collections::Generic::Dictionary<String^, String^>();

	private:
		int currentIndex = 0;

		// функция для получения имени файла из пути
		std::wstring pathFindFileName(std::wstring const& path)
		{
			WCHAR filename[MAX_PATH];
			wcscpy_s(filename, path.c_str());
			PathStripPathW((LPWSTR)filename);
			return std::wstring(filename);
		}

		// функция для получения иконки приложения по имени файла
		HICON GetAppIcon(std::wstring const& filename)
		{
			SHFILEINFOW info;
			SHGetFileInfoW(filename.data(), 0, &info, sizeof(info), SHGFI_ICON | SHGFI_SMALLICON);
			return info.hIcon;
		}

	private: System::Void addBtn_Click(System::Object^ sender, System::EventArgs^ e) {

		try
		{
			// создаем диалоговое окно выбора файла
			OpenFileDialog^ openFileDialog1 = gcnew OpenFileDialog();
			openFileDialog1->Filter = "Applications (*.exe)|*.exe";

			if (openFileDialog1->ShowDialog() == System::Windows::Forms::DialogResult::OK) {
				String^ filename = openFileDialog1->FileName;
				std::wstring wFilename = msclr::interop::marshal_as<std::wstring>(filename);

				LPCWSTR lpwstrFilename = wFilename.c_str();
				std::wstring appName = PathFindFileName(lpwstrFilename);
				String^ appNameString = System::IO::Path::GetFileNameWithoutExtension(gcnew String(appName.c_str()));
				appPaths->Add(appNameString, filename);

				HICON appIcon = GetAppIcon(wFilename);

				Bitmap^ bmp = Bitmap::FromHicon(IntPtr(appIcon));

				// создаем новый элемент ListView с иконкой и названием приложения
				ListViewItem^ item = gcnew ListViewItem(gcnew String(appNameString), 0);

				// создаем словарь для хранения индексов изображений по имени приложения
				System::Collections::Generic::Dictionary<String^, int>^ imageIndexes = nullptr;
				imageIndexes = gcnew System::Collections::Generic::Dictionary<String^, int>();

				// проверяем, есть ли уже изображение для данного приложения
				String^ appNameStr = gcnew String(appName.c_str());
				if (imageIndexes->ContainsKey(appNameStr)) {
					item->ImageIndex = imageIndexes[appNameStr];
				}
				else {
					if (listView1->LargeImageList == nullptr) {
						listView1->LargeImageList = gcnew ImageList();
						listView1->LargeImageList->ImageSize = System::Drawing::Size(32, 32);
					}
					item->ImageIndex = listView1->LargeImageList->Images->Count;
					listView1->LargeImageList->Images->Add(bmp);
					imageIndexes->Add(appNameStr, item->ImageIndex);
				}

				listView1->Items->Add(item);
				DestroyIcon(appIcon);
				currentIndex++;
			}
		}
		catch (Exception^ ex) {
			MessageBox::Show(ex->Message, "Ошибка", MessageBoxButtons::OK, MessageBoxIcon::Error);
		}

	}
	private: System::Void listView1_DoubleClick(System::Object^ sender, System::EventArgs^ e) {

		if (listView1->SelectedItems->Count > 0) {
			auto appPath = appPaths[listView1->SelectedItems[0]->Text];
			Process::Start(appPath);
		}
	}
	private: System::Void MyForm_FormClosing(System::Object^ sender, System::Windows::Forms::FormClosingEventArgs^ e) {

		StreamWriter^ writer = gcnew StreamWriter("appList.txt");

		for each (System::Collections::Generic::KeyValuePair<String^, String^> kvp in appPaths)
		{
			writer->WriteLine(kvp.Key + "|" + kvp.Value);
		}

		writer->Close();
	}
	private: System::Void MyForm_Load(System::Object^ sender, System::EventArgs^ e) {
	
	
		if (File::Exists("appList.txt"))
		{
			StreamReader^ reader = gcnew StreamReader("appList.txt");
	
			String^ line;
			while ((line = reader->ReadLine()) != nullptr)
			{
				array<String^>^ parts = line->Split('|');
				if (parts->Length == 2)
				{
					appPaths->Add(parts[0], parts[1]);
				}
			}
	
			reader->Close();
		}
	
		ImageList^ imageList = gcnew ImageList();
		imageList->ImageSize = System::Drawing::Size(32, 32);
	
		for each (System::Collections::Generic::KeyValuePair<String^, String^> ^ kvp in appPaths)
		{
			String^ appName = kvp->Key;
			String^ appPath = kvp->Value;
			std::wstring appPathStr = msclr::interop::marshal_as<std::wstring>(appPath);
	
			HICON hIcon;
			ExtractIconEx(appPathStr.c_str(), 0, nullptr, &hIcon, 1);
			Bitmap^ bmp = Bitmap::FromHicon(IntPtr(hIcon));
			DestroyIcon(hIcon);
	
			imageList->Images->Add(bmp);
	
			ListViewItem^ item = gcnew ListViewItem(appName, imageList->Images->Count - 1);

			listView1->Items->Add(item);
		}
	
		listView1->LargeImageList = imageList;
	}
	private: System::Void BtnRemove_Click(System::Object^ sender, System::EventArgs^ e) {
	
		if (listView1->SelectedItems->Count > 0) {
			for each (ListViewItem ^ item in listView1->SelectedItems) {
				listView1->Items->Remove(item);
				String^ key = item->Text;
				appPaths->Remove(key);
			}
		}
	}
private: System::Void pictureBox1_Click(System::Object^ sender, System::EventArgs^ e) {
}
};
}

#include <iostream>
#include <windows.h>
#include <string>
#include <vector>
#include <exception>
using namespace std;
void MainMenu();
string OptionsMenu();
void GetRange(int &sRow, int &sCol, int &eRow, int &eCol);

template <typename T>
class Node
{ // Each node has left,right,top and down nodes, data and row and column
public:
    Node<T> *left;
    Node<T> *right;
    Node<T> *up;
    Node<T> *down;
    T data;
    Node(T val)
    {
        left = nullptr;
        right = nullptr;
        up = nullptr;
        down = nullptr;
        data = val;
    }
};

template <typename T>
class MiniExcel
{
private:
    Node<T> *root;    // Root is the top left cell
    Node<T> *current; // Current is the active node
    int rows;
    int cols;

    void SetConsoleColor(int color)
    {
        SetConsoleTextAttribute(GetStdHandle(STD_OUTPUT_HANDLE), color);
    }

public:
    MiniExcel(int row, int col)
    {
        root = nullptr;
        current = nullptr;
        rows = row;
        cols = col;
        CreateDefaultGrid();
    }

    void CreateDefaultGrid()
    {
        root = new Node<T>(" 5 ");
        Node<T> *prevRowStart = root;
        Node<T> *prevColStart = root;
        for (int col = 1; col < cols; col++)
        {
            Node<T> *newcol = new Node<T>(" 5 ");
            newcol->left = prevColStart;
            prevColStart->right = newcol;
            prevColStart = newcol;
        }

        for (int row = 1; row < rows; row++)
        {
            Node<T> *newRow = new Node<T>(" 5 ");
            newRow->up = prevRowStart;
            prevRowStart->down = newRow;
            prevColStart = newRow;
            for (int col = 1; col < cols; col++)
            {
                Node<T> *newcol = new Node<T>(" 5 ");
                if (prevColStart->up != nullptr)
                {
                    newcol->up = prevColStart->up->right;
                    prevColStart->up->right->down = newcol;
                }
                newcol->left = prevColStart;
                prevColStart->right = newcol;
                prevColStart = newcol;
            }
            prevRowStart = newRow;
        }
        current = root;
    }

    void HighlightCurrentCell()
    {
        SetConsoleColor(FOREGROUND_GREEN | FOREGROUND_INTENSITY); // You can choose your own color combination
    }

    void PrintGrid()
    {
        char c = 240;
        system("cls");
        HANDLE H = GetStdHandle(STD_OUTPUT_HANDLE);
        Node<T> *temp = root;
        while (temp != nullptr)
        {
            Node<T> *currentNode = temp;
            while (currentNode != nullptr)
            {
                // cout << "+---";
                cout << "+" << c << c << c << c << c << c;
                currentNode = currentNode->right;
            }
            cout << "+" << endl;

            currentNode = temp;
            while (currentNode != nullptr)
            {
                // cout << "|      |";
                if (currentNode == this->current)
                {

                    cout << "|";
                    SetConsoleTextAttribute(H, 2);
                    if (currentNode->data.length() < 4)
                    {
                        for (int i = currentNode->data.length(); i <= 4; i++)
                        {
                            cout << " ";
                        }
                    }
                    cout << currentNode->data << " ";
                    SetConsoleTextAttribute(H, 7);
                }
                else
                {
                    cout << "|";
                    if (currentNode->data.length() < 4)
                    {
                        for (int i = currentNode->data.length(); i <= 4; i++)
                        {
                            cout << " ";
                        }
                    }
                    cout << currentNode->data << " ";
                }
                currentNode = currentNode->right;
            }
            cout << "|" << endl;

            // cout << "|      |";
            temp = temp->down;
        }

        // Print the bottom border
        Node<T> *current = root;
        while (current != nullptr)
        {
            cout << "+" << c << c << c << c << c << c;
            current = current->right;
        }
        cout << "+" << endl;
    }

    void InsertColumnAtRight()
    {
        Node<T> *node = root;
        while (node->right != nullptr)
        {
            node = node->right;
        }

        while (node != nullptr)
        {
            Node<T> *newNode = new Node<T>("   ");
            node->right = newNode;
            newNode->left = node;
            if (node->up != nullptr)
            {
                newNode->up = node->up->right;
                if (node->up->right != nullptr)
                {
                    node->up->right->down = newNode;
                }
            }
            node = node->down;
        }
        cols++;
    }

    void InsertRowAtEnd()
    {
        Node<T> *downNode = root;
        while (downNode->down != nullptr)
        {
            downNode = downNode->down;
        }

        while (downNode != nullptr)
        {
            Node<T> *newNode = new Node<T>("   ");
            downNode->down = newNode;
            newNode->up = downNode;
            if (downNode->left != nullptr)
            {
                newNode->left = downNode->left->down;
                if (downNode->left->down != nullptr)
                {
                    downNode->left->down->right = newNode;
                }
            }
            downNode = downNode->right;
        }
        rows++;
    }
    void InsertCellByRightShift()
    {
        InsertColumnAtRight();
        Node<T> *last = current;
        while (last->right != nullptr)
        {
            last = last->right;
        }
        Node<T> *chase = last;
        while (chase != current)
        {
            chase->data = chase->left->data;
            chase = chase->left;
        }
        current->data = " ";
        cols++;
    }

    void InsertCellByDownShift()
    {
        InsertRowAtEnd();
        Node<T> *last = current;
        while (last->down != nullptr)
        {
            last = last->down;
        }
        Node<T> *chase = last;
        while (chase != current)
        {
            chase->data = chase->up->data;
            chase = chase->up;
        }
        current->data = " ";
        rows++;
    }

    void DeleteCellByLeftShift()
    {
        if (current->right == nullptr && current->left == nullptr && current->down != nullptr)
        {
            Node<T> *upper = current->up;
            Node<T> *lower = current->down;
            while (upper != nullptr)
            {
                upper->down = lower;
                lower->up = upper;
                upper = upper->right;
                lower = lower->right;
            }
        }
        else if (current->right == nullptr)
        {
            current = nullptr;
        }
        else
        {
            Node<T> *temp = current;
            while (temp->right->right != nullptr)
            {
                temp->data = temp->right->data;
                temp = temp->right;
            }
            if (temp->right->down != nullptr and temp->right->up != nullptr)
            {
                temp->right->down->up = nullptr;
                temp->right->up->down = nullptr;
            }

            temp->right = nullptr;
        }
        if (current->left != nullptr)
        {
            current = current->left;
        }
    }

    void DeleteCellByUpShift()
    {
        if (current->up == nullptr && current->down == nullptr && current->right != nullptr)
        {
            Node<T> *lef = current->left;
            Node<T> *righ = current->right;
            while (lef != nullptr)
            {
                lef->right = righ;
                righ->left = lef;
                righ = righ->down;
                lef = lef->down;
            }
        }
        else if (current->down == nullptr)
        {
            current = nullptr;
        }
        else
        {
            Node<T> *temp = current;
            while (temp->down->down != nullptr)
            {
                temp->data = temp->down->data;
                temp = temp->down;
            }
            if (temp->down->left != nullptr && temp->down->right != nullptr)
            {
                temp->down->left->right = nullptr;
                temp->down->right->left = nullptr;
            }
            temp->down = nullptr;
        }
    }

    // Node<T> *GetUpCell()
    // {
    //     return current->up;
    // }
    // Node<T> *GetDownCell()
    // {
    //     return current->down;
    // }
    // Node<T> *GetRightCell()
    // {
    //     return current->right;
    // }
    // Node<T> *GetLeftCell()
    // {
    //     return current->left;
    // }

    void SetCurrentToLeft()
    {
        current = current->left;
    }
    void SetCurrentToRight()
    {
        current = current->right;
    }
    void SetCurrentToUp()
    {
        current = current->up;
    }
    void SetCurrentToDown()
    {
        current = current->down;
    }

    Node<T> *GetCurrentCell()
    {
        return current;
    }

    void SetCurrentData(T val)
    {
        current->data = val;
    }

    void TakeInput()
    {
        string input;
        getline(cin, input);
        SetCurrentData(input);
        input = "";
    }

    void InserRowAbove()
    {
        Node<T> *node = current;
        while (node->left != nullptr)
        {
            node = node->left;
        }

        while (node != nullptr)
        {
            Node<T> *newNode = new Node<T>("   ");
            newNode->up = node->up;
            if (node->up != nullptr)
            {
                node->up->down = newNode;
            }
            newNode->down = node;
            node->up = newNode;
            if (node->left != nullptr)
            {
                newNode->left = node->left->up;
                if (node->left->up != nullptr)
                {
                    node->left->up->right = newNode;
                }
            }
            node = node->right;
        }
        if (root->up != nullptr)
        {
            root = root->up;
        }
        rows++;
    }

    void InserRowBelow()
    {
        Node<T> *downNode = current;
        while (downNode->left != nullptr)
        {
            downNode = downNode->left;
        }

        while (downNode != nullptr)
        {
            Node<T> *newNode = new Node<T>("   ");
            newNode->down = downNode->down;
            if (downNode->down != nullptr)
            {
                downNode->down->up = newNode;
            }
            newNode->up = downNode;
            downNode->down = newNode;
            if (downNode->left != nullptr)
            {
                newNode->left = downNode->left->down;
                if (downNode->left->down != nullptr)
                {
                    downNode->left->down->right = newNode;
                }
            }
            downNode = downNode->right;
        }
        rows++;
    }

    void InsertColumnToLeft()
    {
        Node<T> *node = current;
        while (node->up != nullptr)
        {
            node = node->up;
        }

        while (node != nullptr)
        {
            Node<T> *newNode = new Node<T>("   ");
            newNode->left = node->left;
            if (node->left != nullptr)
            {
                node->left->right = newNode;
            }
            newNode->right = node;
            node->left = newNode;
            if (node->up != nullptr)
            {
                newNode->up = node->up->left;
                if (node->up->left != nullptr)
                {
                    node->left->up->down = newNode;
                }
            }
            node = node->down;
        }
        // if (current == root)
        if (root->left != nullptr)
        {
            root = root->left;
        }
        cols++;
    }

    void InsertColumnToRight()
    {
        Node<T> *node = current;
        while (node->up != nullptr)
        {
            node = node->up;
        }
        while (node != nullptr)
        {
            Node<T> *newNode = new Node<T>("   ");
            newNode->right = node->right;
            if (node->right != nullptr)
            {
                node->right->left = newNode;
            }
            newNode->left = node;
            node->right = newNode;
            if (node->up != nullptr)
            {
                newNode->up = node->up->right;
                if (node->up->right != nullptr)
                {
                    node->up->right->down = newNode;
                }
            }
            node = node->down;
        }
        cols++;
    }

    void DeleteRow()
    {
        Node<T> *node = current;
        while (node->left != nullptr)
        {
            node = node->left;
        }

        while (node != nullptr)
        {
            if (node->up == nullptr)
            {
                node->down->up = nullptr;
            }
            else if (node->down == nullptr)
            {
                node->up->down = nullptr;
            }
            else
            {
                node->down->up = node->up;
                node->up->down = node->down;
            }
            node = node->right;
        }
        if (current->up == nullptr)
        {
            root = root->down;
            current = current->down;
        }
        else if (current->down == nullptr)
        {
            current = current->up;
        }
        else
        {
            current = current->down;
        }
        rows--;
    }

    void DeleteColumn()
    {
        Node<T> *node = current;
        while (node->up != nullptr)
        {
            node = node->up;
        }

        while (node != nullptr)
        {
            if (node->right == nullptr)
            {
                node->left->right = nullptr;
            }
            else if (node->left == nullptr)
            {
                node->right->left = nullptr;
            }
            else
            {
                node->left->right = node->right;
                node->right->left = node->left;
            }
            node = node->down;
        }
        if (current->left == nullptr)
        {
            root = root->right;
            current = current->right;
        }
        else if (current->down == nullptr)
        {
            current = current->left;
        }
        else
        {
            current = current->left;
        }
        cols--;
    }

    void ClearColumn()
    {
        Node<T> *node = current;
        while (node->up != nullptr)
        {
            node = node->up;
        }
        while (node != nullptr)
        {
            node->data = "";
            node = node->down;
        }
    }

    void ClearRow()
    {
        Node<T> *node = current;
        while (node->left != nullptr)
        {
            node = node->left;
        }
        while (node != nullptr)
        {
            node->data = "";
            node = node->right;
        }
    }

    void gotoxy(int x, int y)
    {
        COORD coordinates;
        coordinates.X = x;
        coordinates.Y = y;
        SetConsoleCursorPosition(GetStdHandle(STD_OUTPUT_HANDLE), coordinates);
    }

    int GetRangeSum(int startRow, int startColumn, int endRow, int endColumn)
    {
        if (startRow < 0 || startRow >= rows || startColumn < 0 || startColumn >= cols ||
            endRow < 0 || endRow >= rows || endColumn < 0 || endColumn >= cols ||
            startRow > endRow || startColumn > endColumn)
        {
            return 0;
        }

        int sum = 0;
        for (int row = startRow; row <= endRow; row++)
        {
            Node<T> *currentRow = root;
            for (int i = 0; i < row; i++)
            {
                currentRow = currentRow->down;
            }

            for (int col = startColumn; col <= endColumn; col++)
            {
                Node<T> *currentCol = currentRow;
                for (int j = 0; j < col; j++)
                {
                    currentCol = currentCol->right;
                }
                int data = stoi(currentCol->data);
                sum += data;
            }
        }

        return sum;
    }

    int GetRangeAverge(int startRow, int startColumn, int endRow, int endColumn)
    {
        int count = 0;
        if (startRow < 0 || startRow >= rows || startColumn < 0 || startColumn >= cols ||
            endRow < 0 || endRow >= rows || endColumn < 0 || endColumn >= cols ||
            startRow > endRow || startColumn > endColumn)
        {
            return 0;
        }

        int sum = 0;

        for (int row = startRow; row <= endRow; row++)
        {
            Node<T> *currentRow = root;
            for (int i = 0; i < row; i++)
            {
                currentRow = currentRow->down;
            }

            for (int col = startColumn; col <= endColumn; col++)
            {
                Node<T> *currentCol = currentRow;
                for (int j = 0; j < col; j++)
                {
                    currentCol = currentCol->right;
                }
                int data = stoi(currentCol->data);
                sum += data;
                count++;
            }
        }

        return sum / count;
    }

    void Sum(int startRow, int startColumn, int endRow, int endColumn)
    {
        int sum = GetRangeSum(startRow, startColumn, endRow, endColumn);
        current->data = sum;
    }

    void Average(int startRow, int startColumn, int endRow, int endColumn)
    {
        int avg = GetRangeAverge(startRow, startColumn, endRow, endColumn);
        current->data = avg;
    }

    void Count(int startRow, int startColumn, int endRow, int endColumn)
    {
        int count = 0;
        // if (startRow < 0 || startRow >= rows || startColumn < 0 || startColumn >= cols ||
        //     endRow < 0 || endRow >= rows || endColumn < 0 || endColumn >= cols ||
        //     startRow > endRow || startColumn > endColumn)
        // {
        //     return T();
        // }

        T sum = T();

        for (int row = startRow; row <= endRow; row++)
        {
            Node<T> *currentRow = root;
            for (int i = 0; i < row; i++)
            {
                currentRow = currentRow->down;
            }

            for (int col = startColumn; col <= endColumn; col++)
            {
                Node<T> *currentCol = currentRow;
                for (int j = 0; j < col; j++)
                {
                    currentCol = currentCol->right;
                }
                count++;
            }
        }

        current->data = to_string(count);
    }

    void Minimum(int startRow, int startColumn, int endRow, int endColumn)
    {
        int min;
        // if (startRow < 0 || startRow >= rows || startColumn < 0 || startColumn >= cols ||
        //     endRow < 0 || endRow >= rows || endColumn < 0 || endColumn >= cols ||
        //     startRow > endRow || startColumn > endColumn)
        // {
        //     min = T();
        // }

        T sum = T();

        for (int row = startRow; row <= endRow; row++)
        {
            Node<T> *currentRow = root;
            for (int i = 0; i < row; i++)
            {
                currentRow = currentRow->down;
            }

            for (int col = startColumn; col <= endColumn; col++)
            {
                Node<T> *currentCol = currentRow;
                for (int j = 0; j < col; j++)
                {
                    currentCol = currentCol->right;
                }
                int data = stoi(currentCol->data);
                if (data < min)
                {
                    min = data;
                }
            }
        }

        current->data = to_string(min);
    }
    void Maximum(int startRow, int startColumn, int endRow, int endColumn)
    {
        int max;
        // if (startRow < 0 || startRow >= rows || startColumn < 0 || startColumn >= cols ||
        //     endRow < 0 || endRow >= rows || endColumn < 0 || endColumn >= cols ||
        //     startRow > endRow || startColumn > endColumn)
        // {
        //     max = T();
        // }

        T sum = T();

        for (int row = startRow; row <= endRow; row++)
        {
            Node<T> *currentRow = root;
            for (int i = 0; i < row; i++)
            {
                currentRow = currentRow->down;
            }

            for (int col = startColumn; col <= endColumn; col++)
            {
                Node<T> *currentCol = currentRow;
                for (int j = 0; j < col; j++)
                {
                    currentCol = currentCol->right;
                }
                int data = stoi(currentCol->data);
                if (data > max)
                {
                    max = data;
                }
            }
        }

        current->data = to_string(max);
    }

    vector<T> cellsData;

    void Copy(int startRow, int startColumn, int endRow, int endColumn)
    {
        char pie = 227;
        string str(1, pie);

        for (int row = startRow; row <= endRow; row++)
        {
            Node<T> *currentRow = root;
            for (int i = 0; i < row; i++)
            {
                currentRow = currentRow->down;
            }

            for (int col = startColumn; col <= endColumn; col++)
            {
                Node<T> *currentCol = currentRow;
                for (int j = 0; j < col; j++)
                {
                    currentCol = currentCol->right;
                }
                cellsData.push_back(currentCol->data);
            }
            cellsData.push_back(str);
        }

        for (int i = 0; i < cellsData.size(); i++)
        {
            cout << cellsData[i];
        }
    }

    void Cut(int startRow, int startColumn, int endRow, int endColumn)
    {
        char pie = 227;
        string str(1, pie);

        T sum = T();

        for (int row = startRow; row <= endRow; row++)
        {
            Node<T> *currentRow = root;
            for (int i = 0; i < row; i++)
            {
                currentRow = currentRow->down;
            }

            for (int col = startColumn; col <= endColumn; col++)
            {
                Node<T> *currentCol = currentRow;
                for (int j = 0; j < col; j++)
                {
                    currentCol = currentCol->right;
                }
                cellsData.push_back(currentCol->data);
                currentCol->data = "   ";
            }
            cellsData.push_back(str);
        }

        for (int i = 0; i < cellsData.size(); i++)
        {
            cout << cellsData[i];
        }
    }

    void Paste()
    {
        char pie = 227;
        string str(1, pie);
        Node<T> *node = current;
        Node<T> *temp = current;
        for (int i = 0; i < cellsData.size() - 1; i++)
        {
            if (cellsData[i] != str)
            {
                if (node->right == nullptr && cellsData[i + 1] != str)
                {
                    InsertColumnAtRight();
                }
                node->data = cellsData[i];
                node = node->right;
            }
            else
            {
                InserRowBelow();
                node = temp->down;
                temp = node;
            }
        }
    }
};
// class RangeStart{
// public:
//     int row;
//     int col;
//     RangeStart(int row,int col)
//     {

//     }

// };
main()
{
    MiniExcel<string> ME(5, 5);
    ME.PrintGrid();
    while (true)
    {
        string option = "";

        if (GetAsyncKeyState(VK_LEFT) && ME.GetCurrentCell()->left != nullptr)
        {
            ME.SetCurrentToLeft();
            ME.PrintGrid();
            MainMenu();
        }
        if (GetAsyncKeyState(VK_RIGHT) && ME.GetCurrentCell()->right != nullptr)
        {
            ME.SetCurrentToRight();
            ME.PrintGrid();
            MainMenu();
        }
        if (GetAsyncKeyState(VK_UP) && ME.GetCurrentCell()->up != nullptr)
        {
            ME.SetCurrentToUp();
            ME.PrintGrid();
            MainMenu();
        }

        if (GetAsyncKeyState(VK_DOWN) && ME.GetCurrentCell()->down != nullptr)
        {
            ME.SetCurrentToDown();
            ME.PrintGrid();
            MainMenu();
        }

        if (GetAsyncKeyState(VK_TAB))
        {
            int startRow, startCol;
            int endRow, endCol;
            option = OptionsMenu();
            if (option == "1")
            {
                ME.InserRowAbove();
            }
            else if (option == "2")
            {
                ME.InserRowBelow();
            }
            else if (option == "3")
            {
                ME.InsertColumnToRight();
            }
            else if (option == "4")
            {
                ME.InsertColumnToLeft();
            }
            else if (option == "5")
            {
                ME.InsertCellByRightShift();
            }
            else if (option == "6")
            {
                ME.InsertCellByDownShift();
            }
            else if (option == "7")
            {
                ME.DeleteCellByLeftShift();
            }
            else if (option == "8")
            {
                ME.DeleteCellByUpShift();
            }
            else if (option == "9")
            {
                ME.DeleteRow();
            }
            else if (option == "10")
            {
                ME.DeleteColumn();
            }
            else if (option == "11")
            {
                ME.ClearRow();
            }
            else if (option == "12")
            {
                ME.ClearColumn();
            }
            else if (option == "13")
            {
                GetRange(startRow, startCol, endRow, endCol);
                ME.Sum(startRow, startCol, endRow, endCol);
            }
            else if (option == "14")
            {
                GetRange(startRow, startCol, endRow, endCol);
                ME.Average(startRow, startCol, endRow, endCol);
            }
            else if (option == "15")
            {
                GetRange(startRow, startCol, endRow, endCol);
                ME.Count(startRow, startCol, endRow, endCol);
            }
            else if (option == "16")
            {
                GetRange(startRow, startCol, endRow, endCol);
                ME.Minimum(startRow, startCol, endRow, endCol);
            }
            else if (option == "17")
            {
                GetRange(startRow, startCol, endRow, endCol);
                ME.Maximum(startRow, startCol, endRow, endCol);
            }
            else if (option == "18")
            {
                GetRange(startRow, startCol, endRow, endCol);
                ME.Copy(startRow, startCol, endRow, endCol);
            }
            else if (option == "19")
            {
                ME.Paste();
            }
            else if (option == "20")
            {
                break;
            }
        }

        else if (GetAsyncKeyState(VK_F1))
        {
            ME.TakeInput();
            ME.PrintGrid();
            MainMenu();
        }
        Sleep(100);
    }
}

void MainMenu()
{
    cout << "***********************************************************************\n";
    cout << "1.\tPress F1 to Enter the data\n";
    cout << "2.\tPress Tab for Options menu\n";
    cout << "+--------------------------------------------+\n";
}
string OptionsMenu()
{
    string option = "";
    cout << "***********************************************************************\n";
    cout << "*                          Options Menu                                \n";
    cout << "***********************************************************************\n";
    cout << "1.\tInsert Row Above the current\n";
    cout << "2.\tInsert Row Below the current\n";
    cout << "3.\tInsert Column Right to the current\n";
    cout << "4.\tInsert Column Left to the current\n";
    cout << "5.\tInsert Cell by Right Shift\n";
    cout << "6.\tInsert Cell by Down Shift\n";
    cout << "7.\tDelete Cell by Left Shift\n";
    cout << "8.\tDelete Cell by Up Shift\n";
    cout << "9.\tDelete Row\n";
    cout << "10.\tDelete Column\n";
    cout << "11.\tClear Row\n";
    cout << "12.\tClear Column\n";
    cout << "13.\tGet Sum\n";
    cout << "14.\tGet Average\n";
    cout << "15.\tCount Total Nodes\n";
    cout << "16.\tFind Minimum value\n";
    cout << "17.\tFind Maximum value\n";
    cout << "18.\tCopy the data\n";
    cout << "19.\tPaste the data\n";
    cout << "20.\tExit\n";
    cout << "+-------------------------------------------------+\n";
    cout << "Enter an option :";
    cin >> option;
    return option;
}

void GetRange(int &sRow, int &sCol, int &eRow, int &eCol)
{
    cout << "Enter starting row :";
    cin >> sRow;
    cout << "Enter starting column :";
    cin >> sCol;
    cout << "Enter ending row :";
    cin >> eRow;
    cout << "Enter endting column :";
    cin >> eCol;
}
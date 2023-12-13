// C++study.cpp : 此文件包含 "main" 函数。程序执行将在此处开始并结束。
//

#include <iostream>
#include "E:/Data(备份)/ZM文件/电子书/C、P、J语言/C++/C++Primer代码/1/Sales_item.h"
using namespace std;

class CPU
{
public:
	CPU();
	~CPU();
	virtual void calculate() = 0;

private:

};

CPU::CPU()
{
}

CPU::~CPU()
{
}

class VideoCard
{
public:
	VideoCard();
	~VideoCard();
	virtual void display() = 0;

private:

};

VideoCard::VideoCard()
{
}

VideoCard::~VideoCard()
{
}


class Memory
{
public:
	Memory();
	~Memory();
	virtual void storage() = 0;
private:

};

Memory::Memory()
{
}

Memory::~Memory()
{
}

class Computer
{
public:
	Computer(CPU* cp, VideoCard* vc, Memory* me);
	~Computer();
	void work()
	{
		cpu->calculate();
		videoCard->display();
		memory->storage();
	}

private:
	CPU* cpu;
	VideoCard* videoCard;
	Memory* memory;
};


Computer::Computer(CPU* cp, VideoCard* vc, Memory* me)
{
	cpu = cp;
	videoCard = vc;
	memory = me;
}

Computer::~Computer()
{
	if (cpu != NULL) {
		delete cpu;
		cpu = NULL;
	}
	if (videoCard != NULL) {
		delete videoCard;
		videoCard = NULL;
	}
	if (memory != NULL) {
		delete memory;
		memory = NULL;
	}
}


class InterCPU : public CPU
{
public:
	InterCPU();
	~InterCPU();
	virtual void calculate()
	{
		cout << "Inter的CPUwork" << endl;
	}

private:

};

InterCPU::InterCPU()
{
}

InterCPU::~InterCPU()
{
}

class InterMemory : public Memory
{
public:
	InterMemory ();
	~InterMemory ();
	virtual void storage()
	{
		cout << "Mem work" << endl;
	}
private:

};

InterMemory ::InterMemory ()
{
}

InterMemory ::~InterMemory ()
{
}

class InterVideoCard : public VideoCard
{
public:
	InterVideoCard();
	~InterVideoCard();
	virtual void display()
	{
		cout << "Video work" << endl;
	}
private:

};

InterVideoCard::InterVideoCard()
{
}

InterVideoCard::~InterVideoCard()
{
}

class LenovoCPU : public CPU
{
public:

	virtual void calculate()
	{
		cout << "Lenovo的CPUwork" << endl;
	}

private:

};

void test01()
{
	CPU* intelCpu = new InterCPU;
	VideoCard* intelvideo = new  InterVideoCard;
	Memory* memory = new InterMemory;

	Computer* computer1 = new Computer(intelCpu, intelvideo, memory);
	computer1->work();
	delete computer1;



}

int main()
{
	test01();

}

// 运行程序: Ctrl + F5 或调试 >“开始执行(不调试)”菜单
// 调试程序: F5 或调试 >“开始调试”菜单

// 入门使用技巧: 
//   1. 使用解决方案资源管理器窗口添加/管理文件
//   2. 使用团队资源管理器窗口连接到源代码管理
//   3. 使用输出窗口查看生成输出和其他消息
//   4. 使用错误列表窗口查看错误
//   5. 转到“项目”>“添加新项”以创建新的代码文件，或转到“项目”>“添加现有项”以将现有代码文件添加到项目
//   6. 将来，若要再次打开此项目，请转到“文件”>“打开”>“项目”并选择 .sln 文件

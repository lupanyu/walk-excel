package main

import (
	"context"
	"log"
	"time"
)

type work struct {
	name   string
	ctx    context.Context
	cancel context.CancelFunc
}

func (t *work) TaskRun() {
	t.ctx, t.cancel = context.WithTimeout(context.Background(), time.Second*5)
	go func(ctx context.Context) {
		if t != nil {
			///
			log.Println(t.name, "in goroutine start")
			time.Sleep(time.Second * time.Duration(200))
		}
	}(t.ctx)
	select {
	case <-t.ctx.Done():
		log.Println(t.name, "goroutine done")
	}
}

func main() {
	//log.Println("main")
	//w := work{"t1", context.Background(), nil}
	//go w.TaskRun()
	//log.Println("t1 goroutine start ok")
	//time.Sleep(10 * time.Second)
	//w2 := work{"t2", context.Background(), nil}
	//go w2.TaskRun()
	//log.Println("t2 goroutine start ok")
	//time.Sleep(2 * time.Second)
	//w2.cancel()
	//time.Sleep(time.Millisecond * 100)
	ch := make(chan int)
	closeCh(ch)
}

func closeCh(ch chan int) {
	ch <- 1
	ch <- 2
	//close(ch)
	for {
		x, ok := <-ch
		if !ok {
			log.Println("ch is closed")
			break
		}
		log.Println(x)
	}
}

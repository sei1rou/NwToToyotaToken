package material

import (
	"errors"
	"log"
)

func failOnError(err error) {
	if err != nil {
		log.Fatal("Error:", err)
	}
}

func doError() error {
	return errors.New("エラーが発生しました。")
}
